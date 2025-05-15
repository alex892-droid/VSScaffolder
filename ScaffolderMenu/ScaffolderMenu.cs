using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;

namespace ScaffolderMenu
{
    internal sealed class ScaffolderMenu
    {
        public static ScaffolderMenu Instance { get; private set; }
        public static readonly Guid CommandSet = new Guid("3c0a1bd0-fce7-48f7-8f22-05fbc20c49d3");
        private readonly AsyncPackage package;
        private DTE2 _dte;

        private ScaffolderMenu(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuItem = new OleMenuCommand(this.Execute, new CommandID(CommandSet, 1));
            commandService.AddCommand(menuItem);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ScaffolderMenu(package, commandService);
            Instance._dte = await package.GetServiceAsync(typeof(DTE)) as DTE2;
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string selectedFolderPath = GetSelectedFolderPath();
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                ShowMessage("Please select a folder in Solution Explorer.");
                return;
            }

            try
            {
                string solutionFolder = Path.GetDirectoryName(_dte.Solution.FullName);
                string scaffoldingFolderPath = Path.Combine(solutionFolder, "Scaffolding");

                if (!Directory.Exists(scaffoldingFolderPath))
                {
                    ShowMessage("Scaffolding folder not found:\n" + scaffoldingFolderPath);
                    return;
                }

                string[] files = Directory.GetFiles(scaffoldingFolderPath, "*.*", SearchOption.TopDirectoryOnly);

                if (files.Length == 0)
                {
                    ShowMessage("No files found in Scaffolding folder.");
                    return;
                }

                int copiedCount = 0;
                foreach (var file in files)
                {
                    string destFilePath = Path.Combine(selectedFolderPath, Path.GetFileName(file));
                    if (!File.Exists(destFilePath))
                    {
                        File.Copy(file, destFilePath);
                        copiedCount++;
                    }
                }

                ShowMessage($"Copied {copiedCount} file(s) to:\n{selectedFolderPath}");
            }
            catch (Exception ex)
            {
                ShowMessage($"Error:\n{ex.Message}", OLEMSGICON.OLEMSGICON_CRITICAL);
            }
        }


        private void ShowMessage(string message, OLEMSGICON icon = OLEMSGICON.OLEMSGICON_INFO)
        {
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                "ScaffolderMenu",
                icon,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private string GetSelectedFolderPath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_dte == null)
                return null;

            Array selectedItems = (Array)_dte.ToolWindows.SolutionExplorer.SelectedItems;
            if (selectedItems == null || selectedItems.Length == 0)
                return null;

            foreach (UIHierarchyItem selectedItem in selectedItems)
            {
                if (selectedItem.Object is ProjectItem projectItem)
                {
                    string filePath = GetProjectItemFullPath(projectItem);
                    if (Directory.Exists(filePath))
                        return filePath;
                }
                else if (selectedItem.Object is Project project)
                {
                    // Handle cases where a project node itself is selected
                    string projectPath = project.FullName;
                    if (!string.IsNullOrEmpty(projectPath))
                        return Path.GetDirectoryName(projectPath);
                }
            }

            return null;
        }

        private string GetProjectItemFullPath(ProjectItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (item == null)
                return null;

            try
            {
                if (item.Properties == null)
                    return null;

                foreach (Property prop in item.Properties)
                {
                    if (prop != null && prop.Name == "FullPath")
                    {
                        return prop.Value as string;
                    }
                }
            }
            catch (Exception)
            {
                // Ignored (sometimes properties throw exceptions)
            }

            return null;
        }
    }
}
