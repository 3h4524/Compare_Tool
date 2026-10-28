# Compare Tool

Windows WPF MVP để so sánh file Word trong hai folder Old/New.

## Chạy

```powershell
dotnet run
```

## Đã có

- Chọn Old, New và Output folder.
- Scan đệ quy các file `.docx`, `.docm`, `.doc`, `.dotx`, `.dotm`, `.rtf`.
- Ghép file theo tên không phân biệt hoa thường.
- So sánh SHA-256.
- Hiển thị `Trùng`, `Khác`, `Chỉ có Old`, `Chỉ có New`.
- Move theo nhóm vào Output, không ghi đè file đích.
- Mỗi lần Move tạo một batch mới: `Batch1`, `Batch2`, ...; file khác nhau được gom theo từng cặp trong `BatchN/Different/<file-name>/Old.<ext>` và `New.<ext>`.
- Preview cơ bản hai cột cho `.docx`/`.docm`.

## Tích hợp Aspose sau này

Đặt `Aspose.Words.dll` cạnh file `.exe` và đặt một trong các license sau cạnh file `.exe`:

```text
Aspose.Total.NET.lic
Aspose.Words.lic
AsposeLicense.lic
```

Hoặc đặt biến môi trường `ASPOSE_WORDS_LICENSE` trỏ tới file license. Khi đủ cả DLL và license, nút compare sẽ dùng Aspose; nếu thiếu, app không gọi Aspose và dùng preview/Word Interop fallback.

API dự kiến:

```csharp
var oldDoc = new Aspose.Words.Document(oldPath);
var newDoc = new Aspose.Words.Document(newPath);
var options = new Aspose.Words.Comparing.CompareOptions
{
    Granularity = Aspose.Words.Comparing.Granularity.Word,
    IgnoreFormatting = false,
    IgnoreHeadersAndFooters = false,
    IgnoreTables = false
};
oldDoc.Compare(newDoc, "CompareTool", DateTime.Now, options);
var different = oldDoc.Revisions.Count > 0;
oldDoc.Save(resultPath);
```

Có thể mở `resultPath` bằng Word để review Track Changes. Word Interop được cấu hình compare toàn bộ: text, formatting, case, whitespace, tables, headers/footers, footnotes, textboxes, fields, comments và moves. License không được commit vào source code.
