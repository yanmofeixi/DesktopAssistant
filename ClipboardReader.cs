using System.Drawing.Imaging;

namespace DesktopAssistant;

/// <summary>Reads the logged-in desktop clipboard; the SSH session has its own, unrelated clipboard.</summary>
public static class ClipboardReader
{
    public sealed record Snapshot(string? Text, string? Html, string[]? Files, Size? ImageSize, byte[]? ImagePng);

    public static Snapshot Read(bool includeImageBytes)
    {
        Snapshot? snapshot = null;
        Exception? failure = null;
        // Clipboard access requires an STA thread.
        var thread = new Thread(() =>
        {
            try { snapshot = ReadOnSta(includeImageBytes); }
            catch (Exception ex) { failure = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (failure != null) throw new CommandException($"读取剪贴板失败：{failure.Message}", 500);
        return snapshot!;
    }

    private static Snapshot ReadOnSta(bool includeImageBytes)
    {
        string? text = Clipboard.ContainsText(TextDataFormat.UnicodeText) ? Clipboard.GetText(TextDataFormat.UnicodeText) : null;
        string? html = Clipboard.ContainsText(TextDataFormat.Html) ? StripHtmlClipboardHeader(Clipboard.GetText(TextDataFormat.Html)) : null;
        string[]? files = Clipboard.ContainsFileDropList() ? Clipboard.GetFileDropList().Cast<string>().ToArray() : null;
        Size? imageSize = null;
        byte[]? png = null;
        if (Clipboard.ContainsImage() && Clipboard.GetImage() is { } image)
        {
            using (image)
            {
                imageSize = image.Size;
                if (includeImageBytes)
                {
                    using var stream = new MemoryStream();
                    image.Save(stream, ImageFormat.Png);
                    png = stream.ToArray();
                }
            }
        }
        return new Snapshot(text, html, files, imageSize, png);
    }

    /// <summary>CF_HTML prefixes the markup with "Version:/StartHTML:..." lines whose offsets count UTF-8 bytes.</summary>
    private static string StripHtmlClipboardHeader(string raw)
    {
        int start = raw.IndexOf('<');
        return start < 0 ? raw : raw[start..];
    }
}
