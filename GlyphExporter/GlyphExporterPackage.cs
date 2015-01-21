using System;
using System.ComponentModel.Design;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace MadsKristensen.GlyphExporter
{

	[PackageRegistration(UseManagedResourcesOnly = true)]
	[InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
	[ProvideMenuResource("Menus.ctmenu", 1)]
	[Guid(GuidList.guidGlyphExporterPkgString)]
	public sealed class GlyphExporterPackage : Package
	{
		private IGlyphService _glyphService;
		private IVsImageService2 _imageService;

		protected override void Initialize()
		{
			base.Initialize();

			IComponentModel model = (IComponentModel)this.GetService(typeof(SComponentModel));
			_glyphService = model.GetService<IGlyphService>();
			_imageService = GetService(typeof(SVsImageService)) as IVsImageService2;

			OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

			CommandID cmdGlyphId = new CommandID(GuidList.guidGlyphExporterCmdSet, (int)PkgCmdIDList.cmdidGlyph);
			MenuCommand menuGlyph = new MenuCommand(ButtonClicked, cmdGlyphId);
			mcs.AddCommand(menuGlyph);
		}

		private void SaveImagesToDisk(string folder)
		{
			PropertyInfo[] monikers = typeof(KnownMonikers).GetProperties(BindingFlags.Static | BindingFlags.Public);

			ImageAttributes imageAttributes = new ImageAttributes();
			imageAttributes.Flags = (uint)_ImageAttributesFlags.IAF_RequiredFlags;
			imageAttributes.ImageType = (uint)_UIImageType.IT_Bitmap;
			imageAttributes.Format = (uint)_UIDataFormat.DF_WPF;
			imageAttributes.LogicalHeight = 16;
			imageAttributes.LogicalWidth = 16;
			imageAttributes.StructSize = Marshal.SizeOf(typeof(ImageAttributes));

			foreach (var monikerName in monikers)
			{
				ImageMoniker moniker = (ImageMoniker)monikerName.GetValue(null, null);
				IVsUIObject result = _imageService.GetImage(moniker, imageAttributes);

				Object data;
				result.get_Data(out data);

				if (data == null)
					continue;

				BitmapSource glyph = data as BitmapSource;

				string fileName = Path.Combine(folder, monikerName.Name + ".png");

				SaveBitmapToDisk(glyph, fileName);
			}
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

		private static string GetFolder()
		{
			using (FolderBrowserDialog dialog = new FolderBrowserDialog())
			{
				dialog.RootFolder = Environment.SpecialFolder.DesktopDirectory;
				dialog.ShowDialog();

				return dialog.SelectedPath;
			}
		}

		private void SaveGlyphsToDisk(string folder)
		{
			Array groups = Enum.GetValues(typeof(StandardGlyphGroup));
			Array items = Enum.GetValues(typeof(StandardGlyphItem));

			foreach (var groupName in groups)
				foreach (var itemName in items)
				{
					StandardGlyphGroup group = (StandardGlyphGroup)groupName;
					StandardGlyphItem item = (StandardGlyphItem)itemName;

					BitmapSource glyph = _glyphService.GetGlyph(group, item) as BitmapSource;

					if (glyph == null)
						continue;

					string fileName = Path.Combine(folder, group.ToString(), item.ToString() + ".png");

					SaveBitmapToDisk(glyph, fileName);
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
	}
}
