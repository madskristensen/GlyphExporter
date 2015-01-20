using System;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;

namespace MadsKristensen.GlyphExporter
{

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidGlyphExporterPkgString)]
    public sealed class GlyphExporterPackage : Package
    {
        internal IGlyphService _glyphService { get; set; }

        protected override void Initialize()
        {
            base.Initialize();

            IComponentModel model = (IComponentModel)this.GetService(typeof(SComponentModel));
            _glyphService = model.GetService<IGlyphService>();

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                CommandID menuCommandID = new CommandID(GuidList.guidGlyphExporterCmdSet, (int)PkgCmdIDList.cmdidMyCommand);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.RootFolder = Environment.SpecialFolder.DesktopDirectory;
                dialog.ShowDialog();

                string folder = dialog.SelectedPath;

                if (!string.IsNullOrEmpty(folder) && Directory.Exists(folder))
                {
                    SaveGlyphsToDisk(folder);
                }
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
}
