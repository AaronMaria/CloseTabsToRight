using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Platform.WindowManagement;
using Microsoft.VisualStudio.PlatformUI.Shell;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Threading;
using Task = System.Threading.Tasks.Task;
using static CloseTabsToRight.Helpers.WindowFrameHelpers;
using static CloseTabsToRight.Helpers.DocumentHelpers;

namespace CloseTabsToRight.Commands {
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CloseTabsToRightCommand {
        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package, DTE2 dte) {
            if (package == null) {
                throw new ArgumentNullException(nameof(package));
            }
            var commandService = (IMenuCommandService)await package.GetServiceAsync(typeof(IMenuCommandService));
            if (commandService is OleMenuCommandService) {
                var id = new CommandID(PackageGuids.GuidCommandPackageCmdSet, PackageIds.CloseTabsToRightCommandId);
                var command = new OleMenuCommand((s, e) => CloseTabsToRight(package, dte), id);
                command.BeforeQueryStatus += (sender, e) => BeforeQueryStatus(sender, e, package, dte);
                commandService.AddCommand(command);
            }
        }

        private static void BeforeQueryStatus(object sender, EventArgs e, AsyncPackage package, DTE2 dte) {
            var button = (OleMenuCommand)sender;

            var vsWindowFrames = GetVsWindowFrames(package).ToList();
            var activeFrame = GetActiveWindowFrame(vsWindowFrames, dte);
            var docGroup = GetDocumentGroup(activeFrame);

            var docViewsToRight = GetDocumentViewsToRight(activeFrame, docGroup);

            button.Enabled = docViewsToRight.Any();
        }

        private static void CloseTabsToRight(AsyncPackage package, DTE2 dte) {
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
            var foundActive = false;
            foreach(var name in documentViews.Select(documentView => CleanDocumentViewName(documentView.Name))) {
                if(!foundActive) {
                    if(name == viewMoniker) {
                        foundActive = true;

                    }

                    // Skip over documents until we have found the first one after the active
                    continue;
                }

                var frame = windowFramesDict[name];
                if(frame != null)
                    framesToClose.Add(frame);
            }

            foreach(var frame in framesToClose) {
                frame.CloseFrame(__FRAMECLOSE.FRAMECLOSE_PromptSave);
            }
        }

        private static IEnumerable<DocumentView> GetDocumentViewsToRight(WindowFrame activeWindowFrame, DocumentGroup docGroup) {
            var docViewsToRight = new List<DocumentView>();
            var viewMoniker = activeWindowFrame.FrameMoniker.ViewMoniker;
            var documentViews = docGroup.Children.Where(c => c != null && c.GetType() == typeof(DocumentView)).Select(c => c as DocumentView);
            var foundActive = false;

            foreach(var documentView in documentViews) {
                var name = CleanDocumentViewName(documentView.Name);
                if(!foundActive) {
                    if(name == viewMoniker) {
                        foundActive = true;

                    }

                    // Skip over documents until we have found the first one after the active
                    continue;
                }

                docViewsToRight.Add(documentView);
            }

            return docViewsToRight;
        }
    }
}
