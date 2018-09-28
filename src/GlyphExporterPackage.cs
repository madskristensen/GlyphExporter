using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace MadsKristensen.GlyphExporter
{
	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
	[ProvideMenuResource("Menus.ctmenu", 1)]
	[Guid(PackageGuids.guidGlyphExporterPkgString)]
	public sealed class GlyphExporterPackage : AsyncPackage
    {
		private IGlyphService _glyphService;
		private IVsImageService2 _imageService;

		protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
		{
            await base.InitializeAsync(cancellationToken, progress);

            // When initialized asynchronously, we *may* be on a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            // Otherwise, remove the switch to the UI thread if you don't need it.
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var model = (IComponentModel)await GetServiceAsync(typeof(SComponentModel));
            Assumes.Present(model);
            _glyphService = model.GetService<IGlyphService>();

            _imageService = await GetServiceAsync(typeof(SVsImageService)) as IVsImageService2;
            Assumes.Present(_imageService);

            var mcs = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Assumes.Present(mcs);

            var cmdGlyphId = new CommandID(PackageGuids.guidGlyphExporterCmdSet, PackageIds.cmdidGlyph);
			var menuGlyph = new MenuCommand(ButtonClicked, cmdGlyphId);
			mcs.AddCommand(menuGlyph);
		}

		private void ButtonClicked(object sender, EventArgs e)
		{
			string folder = GetFolder();

			if (!string.IsNullOrEmpty(folder) && Directory.Exists(folder))
			{
				SaveImagesToDisk(Path.Combine(folder, "images"));
				SaveGlyphsToDisk(Path.Combine(folder, "glyphs"));
			}
		}

		private void SaveImagesToDisk(string folder)
		{
            ThreadHelper.ThrowIfNotOnUIThread();

            PropertyInfo[] monikers = typeof(KnownMonikers).GetProperties(BindingFlags.Static | BindingFlags.Public);

            var imageAttributes = new ImageAttributes
            {
                Flags = (uint)_ImageAttributesFlags.IAF_RequiredFlags,
                ImageType = (uint)_UIImageType.IT_Bitmap,
                Format = (uint)_UIDataFormat.DF_WPF,
                LogicalHeight = 16,
                LogicalWidth = 16,
                StructSize = Marshal.SizeOf(typeof(ImageAttributes))
            };

            WriteableBitmap sprite = null;
			int count = 0;
			char letter = ' ';

			foreach (PropertyInfo monikerName in monikers)
			{
				var moniker = (ImageMoniker)monikerName.GetValue(null, null);
				IVsUIObject result = _imageService.GetImage(moniker, imageAttributes);

				if (monikerName.Name[0] != letter)
				{
					if (sprite != null)
					{
						sprite.Unlock();
						SaveBitmapToDisk(sprite, Path.Combine(folder, "_sprites", letter + ".png"));
					}

					int items = monikers.Count(m => m.Name[0] == monikerName.Name[0]);
					sprite = new WriteableBitmap(16, 16 * (items), 96, 96, PixelFormats.Pbgra32, null);
					sprite.Lock();
					letter = monikerName.Name[0];
					count = 0;
				}

                result.get_Data(out object data);

                if (data == null)
					continue;

				var glyph = data as BitmapSource;
				string fileName = Path.Combine(folder, monikerName.Name + ".png");

				int stride = glyph.PixelWidth * (glyph.Format.BitsPerPixel / 8);
				byte[] buffer = new byte[stride * glyph.PixelHeight];
				glyph.CopyPixels(buffer, stride, 0);

				sprite.WritePixels(new Int32Rect(0, count, glyph.PixelWidth, glyph.PixelHeight), buffer, stride, 0);
				count += 16;

				SaveBitmapToDisk(glyph, fileName);
			}

			// Save the last image sprite to disk
			sprite.Unlock();
			SaveBitmapToDisk(sprite, Path.Combine(folder, "_sprites", letter + ".png"));
		}


		private void SaveGlyphsToDisk(string folder)
		{
			Array groups = Enum.GetValues(typeof(StandardGlyphGroup));
			Array items = Enum.GetValues(typeof(StandardGlyphItem));

			foreach (object groupName in groups)
			{
				int count = 0;
				string glyphFolder = Path.Combine(folder, groupName.ToString());
				var sprite = new WriteableBitmap(16, 16 * (items.Length), 96, 96, PixelFormats.Pbgra32, null);
				sprite.Lock();

				foreach (object itemName in items)
				{
					var group = (StandardGlyphGroup)groupName;
					var item = (StandardGlyphItem)itemName;

                    if (!(_glyphService.GetGlyph(group, item) is BitmapSource glyph))
                        continue;

                    string fileName = Path.Combine(folder, group.ToString(), item.ToString() + ".png");

					SaveBitmapToDisk(glyph, fileName);

					int stride = glyph.PixelWidth * (glyph.Format.BitsPerPixel / 8);
					byte[] data = new byte[stride * glyph.PixelHeight];
					glyph.CopyPixels(data, stride, 0);

					sprite.WritePixels(
						new Int32Rect(0, count, glyph.PixelWidth, glyph.PixelHeight),
						data, stride, 0);

					count += 16;
				}

				sprite.Unlock();
				SaveBitmapToDisk(sprite, Path.Combine(folder, "_sprites", groupName + ".png"));
			}
		}

		private static void SaveBitmapToDisk(BitmapSource glyph, string fileName)
		{
			Directory.CreateDirectory(Path.GetDirectoryName(fileName));

			using (var fileStream = new FileStream(fileName, FileMode.Create))
			{
				BitmapEncoder encoder = new PngBitmapEncoder();
				encoder.Frames.Add(BitmapFrame.Create(glyph));
				encoder.Save(fileStream);
			}
		}

		private static string GetFolder()
		{
			using (var dialog = new FolderBrowserDialog())
			{
				dialog.RootFolder = Environment.SpecialFolder.DesktopDirectory;
				dialog.ShowDialog();

				return dialog.SelectedPath;
			}
		}
	}
}
