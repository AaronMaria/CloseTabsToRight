using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Platform.WindowManagement;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using static CloseTabsToRight.Helpers.WindowFrameHelpers;
using static CloseTabsToRight.Helpers.DocumentHelpers;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.PlatformUI.Shell;

namespace CloseTabsToRight.Commands {
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CloseTabsToLeftCommand {

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package, DTE2 dte) {
            if(package == null) {
                throw new ArgumentNullException(nameof(package));
            }
            var commandService = (IMenuCommandService)await package.GetServiceAsync(typeof(IMenuCommandService));
            if (commandService is OleMenuCommandService) {
                var id = new CommandID(PackageGuids.GuidCommandPackageCmdSet, PackageIds.CloseTabsToLeftCommandId);
                var command = new OleMenuCommand((s, e) => CloseTabsToLeft(package, dte), id);
                command.BeforeQueryStatus += (sender, e) => BeforeQueryStatus(sender, e, package, dte);
                commandService.AddCommand(command);
            }
        }

        private static void BeforeQueryStatus(object sender, EventArgs e, AsyncPackage package, DTE2 dte) {
            var button = (OleMenuCommand)sender;

            var vsWindowFrames = GetVsWindowFrames(package).ToList();
            var activeFrame = GetActiveWindowFrame(vsWindowFrames, dte);
            var docGroup = GetDocumentGroup(activeFrame);

            var docViewsToRight = GetDocumentViewsToLeft(activeFrame, docGroup);
            button.Enabled = docViewsToRight.Any();
        }
        private static IEnumerable<DocumentView> GetDocumentViewsToLeft(WindowFrame activeWindowFrame, DocumentGroup docGroup) {
            var docViewsToRight = new List<DocumentView>();
            var viewMoniker = activeWindowFrame.FrameMoniker.ViewMoniker;
            var documentViews = docGroup.Children.Where(c => c != null && c.GetType() == typeof(DocumentView)).Select(c => c as DocumentView);
            var foundActive = false;

            foreach (var documentView in documentViews) {
                var name = CleanDocumentViewName(documentView.Name);
                if (!foundActive) {
                    if (name == viewMoniker) {
                        foundActive = true;
                    } else {
                        docViewsToRight.Add(documentView);
                    }
                }

            }

            return docViewsToRight;
        }

        private static void CloseTabsToLeft(AsyncPackage package, DTE2 dte) {
            var vsWindowFrames = GetVsWindowFrames(package).ToList();
            var windowFrames = vsWindowFrames.Select(vsWindowFrame => vsWindowFrame as WindowFrame);
            var activeFrame = GetActiveWindowFrame(vsWindowFrames, dte);

            var windowFrame = activeFrame;
            if(windowFrame == null)
                return;

            var windowFramesDict = windowFrames.ToDictionary(frame => frame.FrameMoniker.ViewMoniker); 
            
            var docGroup = GetDocumentGroup(windowFrame);
            var viewMoniker = windowFrame.FrameMoniker.ViewMoniker;
            var documentViews = docGroup.Children.Where(c => c != null && c.GetType() == typeof(DocumentView)).Select(c => c as DocumentView);

            var framesToClose = new List<WindowFrame>();
            foreach(var name in documentViews.Select(documentView => CleanDocumentViewName(documentView.Name))) {
                if(name == viewMoniker) {
                    // We found the active tab. No need to continue
                    break;
                }

                var frame = windowFramesDict[name];
                if(frame != null)
                    framesToClose.Add(frame);
            }

            foreach(var frame in framesToClose) {
                frame.CloseFrame(__FRAMECLOSE.FRAMECLOSE_PromptSave);
            }
        }
    }
}
