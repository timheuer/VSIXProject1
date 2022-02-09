using Microsoft.VisualStudio.Text;

namespace VSIXProject1
{
    [Command(PackageIds.MyCommand)]
    internal sealed class MyCommand : BaseCommand<MyCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await Package.JoinableTaskFactory.SwitchToMainThreadAsync();
            DocumentView docView = await VS.Documents.GetActiveDocumentViewAsync();
            if (docView?.TextView == null) return;
            SnapshotPoint position = docView.TextView.Caret.Position.BufferPosition;
            docView.TextBuffer?.Insert(position, Guid.NewGuid().ToString());
        }
    }
}
