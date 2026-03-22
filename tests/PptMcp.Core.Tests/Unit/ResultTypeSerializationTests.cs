using System.Text.Json;
using System.Globalization;
using PptMcp.Core.Models;
using Xunit;

namespace PptMcp.Core.Tests.Unit;

/// <summary>
/// Validates JSON serialization behavior of result types,
/// ensuring null properties are omitted and camelCase naming works correctly.
/// </summary>
public class ResultTypeSerializationTests
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    [Fact]
    public void OperationResult_Success_OmitsNullFields()
    {
        var result = new OperationResult { Success = true, Action = "create", Message = "Done" };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"success\":true", json);
        Assert.Contains("\"action\":\"create\"", json);
        Assert.DoesNotContain("errorMessage", json);
        Assert.DoesNotContain("filePath", json);
    }

    [Fact]
    public void OperationResult_Failure_IncludesErrorMessage()
    {
        var result = new OperationResult { Success = false, ErrorMessage = "Not found" };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"success\":false", json);
        Assert.Contains("\"errorMessage\":\"Not found\"", json);
    }

    [Fact]
    public void SlideListResult_WithSlides_SerializesCorrectly()
    {
        var result = new SlideListResult
        {
            Success = true,
            Slides =
            [
                new SlideInfo
                {
                    SlideIndex = 1,
                    SlideNumber = 1,
                    SlideId = "256",
                    LayoutName = "Title Slide",
                    MasterName = "Office Theme",
                    ShapeCount = 3
                }
            ]
        };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"slideIndex\":1", json);
        Assert.Contains("\"layoutName\":\"Title Slide\"", json);
        Assert.Contains("\"shapeCount\":3", json);
    }

    [Fact]
    public void ShapeInfo_NullOptionalFields_AreOmitted()
    {
        var info = new ShapeInfo
        {
            ShapeId = 1,
            Name = "Rectangle 1",
            ShapeType = "AutoShape",
            Width = 100f,
            Height = 50f
        };
        var json = JsonSerializer.Serialize(info, JsonOptions);

        Assert.Contains("\"name\":\"Rectangle 1\"", json);
        Assert.DoesNotContain("\"text\":", json);
        Assert.DoesNotContain("\"alternativeText\":", json);
        Assert.DoesNotContain("\"placeholderType\":", json);
        Assert.DoesNotContain("\"groupItems\":", json);
    }

    [Fact]
    public void TextResult_WithParagraphs_SerializesNestedStructure()
    {
        var result = new TextResult
        {
            Success = true,
            ShapeId = 1,
            ShapeName = "Title 1",
            Text = "Hello World",
            Paragraphs =
            [
                new TextParagraphInfo
                {
                    Index = 0,
                    Text = "Hello World",
                    Runs =
                    [
                        new TextRunInfo { Text = "Hello ", Bold = true, FontSize = 24f },
                        new TextRunInfo { Text = "World", Italic = true }
                    ]
                }
            ]
        };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"bold\":true", json);
        Assert.Contains("\"fontSize\":24", json);
        Assert.Contains("\"italic\":true", json);
    }

    [Fact]
    public void OperationResult_RoundTrip_PreservesAllFields()
    {
        var original = new OperationResult
        {
            Success = true,
            Action = "delete",
            Message = "Deleted slide 3",
            FilePath = @"C:\test\pres.pptx"
        };

        var json = JsonSerializer.Serialize(original, JsonOptions);
        var deserialized = JsonSerializer.Deserialize<OperationResult>(json, JsonOptions);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Success, deserialized.Success);
        Assert.Equal(original.Action, deserialized.Action);
        Assert.Equal(original.Message, deserialized.Message);
        Assert.Equal(original.FilePath, deserialized.FilePath);
    }

    [Fact]
    public void DocumentPropertyResult_AllNulls_MinimalJson()
    {
        var result = new DocumentPropertyResult { Success = true };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        // Should only have success, no null property fields
        Assert.Contains("\"success\":true", json);
        Assert.DoesNotContain("\"title\":", json);
        Assert.DoesNotContain("\"author\":", json);
        Assert.DoesNotContain("\"subject\":", json);
    }

    [Fact]
    public void HyperlinkInfo_ConditionalSerialization_WhenWritingDefault()
    {
        var info = new HyperlinkInfo
        {
            Index = 1,
            Address = "https://example.com"
        };
        var json = JsonSerializer.Serialize(info, JsonOptions);

        Assert.Contains("\"address\":\"https://example.com\"", json);
        // SlideIndex = 0 should be omitted (WhenWritingDefault)
        Assert.DoesNotContain("\"slideIndex\":", json);
    }

    [Fact]
    public void ArchetypeDetailResult_ObservedExamples_RemainSanitized()
    {
        var result = new ArchetypeDetailResult
        {
            Success = true,
            Id = "framework",
            Name = "Framework",
            ObservedExampleSlides = ["ref-sample123"],
            ObservedExamples =
            [
                new ReferenceSlideInfo
                {
                    Id = "ref-sample123",
                    ArchetypeId = "framework",
                    SubArchetypeId = "matrix-grid",
                    Rationale = "Sample rationale."
                }
            ],
            ObservedSubtypes =
            [
                new ReferenceSubtypeInfo
                {
                    SubArchetypeId = "matrix-grid",
                    Count = 1,
                    ExampleSlides = ["ref-sample123"],
                    ExampleDetails =
                    [
                        new ReferenceSlideInfo
                        {
                            Id = "ref-sample123",
                            ArchetypeId = "framework",
                            SubArchetypeId = "matrix-grid",
                            Rationale = "Sample rationale."
                        }
                    ]
                }
            ]
        };

        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"observedExampleSlides\":[\"ref-sample123\"]", json);
        Assert.Contains("\"observedExamples\":[", json);
        Assert.Contains("\"exampleDetails\":[", json);
        Assert.DoesNotContain("\"sourceName\":", json);
        Assert.DoesNotContain("\"sourceImage\":", json);
        Assert.DoesNotContain("\"batchId\":", json);
        Assert.DoesNotContain("\"deckKey\":", json);
    }

    [Fact]
    public void MailListResult_WithOutlookMessages_SerializesExpectedFields()
    {
        var result = new MailListResult
        {
            Success = true,
            FolderName = "Drafts",
            Query = "copilot",
            TotalItemCount = 2,
            ReturnedCount = 1,
            Messages =
            [
                new MailSummaryInfo
                {
                    EntryId = "entry-1",
                    StoreId = "store-1",
                    Subject = "Copilot draft",
                    SenderName = "Torsten",
                    To = "team@example.com",
                    BodyPreview = "Preview text",
                    Categories = ["Copilot", "Follow Up"],
                    Unread = false,
                    IsDraft = true,
                    Importance = 2,
                    AttachmentCount = 0,
                    SentOn = DateTimeOffset.Parse("2026-03-21T22:00:00+00:00", CultureInfo.InvariantCulture)
                }
            ]
        };

        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"folderName\":\"Drafts\"", json);
        Assert.Contains("\"query\":\"copilot\"", json);
        Assert.Contains("\"messages\":[", json);
        Assert.Contains("\"categories\":[\"Copilot\",\"Follow Up\"]", json);
        Assert.Contains("\"isDraft\":true", json);
        Assert.Contains("\"attachmentCount\":0", json);
        Assert.Contains("\"sentOn\":\"2026-03-21T22:00:00+00:00\"", json);
    }

    [Fact]
    public void MailSendResult_NullOptionalFields_AreOmitted()
    {
        var result = new MailSendResult
        {
            Success = false,
            Sent = false,
            ErrorMessage = "The selected Outlook mail item has already been sent."
        };

        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"sent\":false", json);
        Assert.Contains("\"errorMessage\":\"The selected Outlook mail item has already been sent.\"", json);
        Assert.DoesNotContain("\"entryId\":", json);
        Assert.DoesNotContain("\"storeId\":", json);
        Assert.DoesNotContain("\"subject\":", json);
        Assert.DoesNotContain("\"message\":", json);
    }

    [Fact]
    public void MailMutationResult_WithFolderState_RoundTrips()
    {
        var original = new MailMutationResult
        {
            Success = true,
            EntryId = "entry-1",
            StoreId = "store-1",
            Subject = "Copilot move",
            FolderName = "Deleted Items",
            FolderPath = "\\Mailbox - Test\\Deleted Items",
            Categories = ["Copilot"],
            Moved = true,
            Read = false,
            Message = "Moved Outlook mail item."
        };

        var json = JsonSerializer.Serialize(original, JsonOptions);
        var deserialized = JsonSerializer.Deserialize<MailMutationResult>(json, JsonOptions);

        Assert.Contains("\"folderName\":\"Deleted Items\"", json);
        Assert.Contains("\"categories\":[\"Copilot\"]", json);
        Assert.Contains("\"moved\":true", json);
        Assert.Contains("\"read\":false", json);
        Assert.NotNull(deserialized);
        Assert.Equal(original.FolderPath, deserialized.FolderPath);
        Assert.Equal(original.Moved, deserialized.Moved);
        Assert.Equal(original.Read, deserialized.Read);
    }

    [Fact]
    public void CalendarListResult_WithAppointments_SerializesExpectedFields()
    {
        var result = new CalendarListResult
        {
            Success = true,
            FolderName = "Calendar",
            TotalItemCount = 3,
            ReturnedCount = 1,
            Appointments =
            [
                new CalendarSummaryInfo
                {
                    EntryId = "appt-1",
                    StoreId = "store-1",
                    Subject = "Copilot review",
                    Location = "Room A",
                    Start = DateTimeOffset.Parse("2026-03-22T10:00:00+00:00", CultureInfo.InvariantCulture),
                    End = DateTimeOffset.Parse("2026-03-22T10:30:00+00:00", CultureInfo.InvariantCulture),
                    ReminderSet = true,
                    BusyStatus = 2
                }
            ]
        };

        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"folderName\":\"Calendar\"", json);
        Assert.Contains("\"appointments\":[", json);
        Assert.Contains("\"location\":\"Room A\"", json);
        Assert.Contains("\"reminderSet\":true", json);
    }

    [Fact]
    public void CalendarAppointmentResult_OmitsNullOptionalFields()
    {
        var result = new CalendarAppointmentResult
        {
            Success = true,
            Saved = true,
            Displayed = false,
            AllDay = false
        };

        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"saved\":true", json);
        Assert.Contains("\"displayed\":false", json);
        Assert.DoesNotContain("\"entryId\":", json);
        Assert.DoesNotContain("\"subject\":", json);
        Assert.DoesNotContain("\"message\":", json);
    }

    [Fact]
    public void CalendarMutationResult_WithState_RoundTrips()
    {
        var original = new CalendarMutationResult
        {
            Success = true,
            EntryId = "appt-2",
            StoreId = "store-2",
            Subject = "Updated review",
            Location = "Room B",
            Start = DateTimeOffset.Parse("2026-03-22T11:00:00+00:00", CultureInfo.InvariantCulture),
            End = DateTimeOffset.Parse("2026-03-22T11:30:00+00:00", CultureInfo.InvariantCulture),
            Updated = true,
            Deleted = false,
            Message = "Updated Outlook appointment."
        };

        var json = JsonSerializer.Serialize(original, JsonOptions);
        var deserialized = JsonSerializer.Deserialize<CalendarMutationResult>(json, JsonOptions);

        Assert.Contains("\"updated\":true", json);
        Assert.Contains("\"deleted\":false", json);
        Assert.Contains("\"location\":\"Room B\"", json);
        Assert.NotNull(deserialized);
        Assert.Equal(original.Subject, deserialized.Subject);
        Assert.Equal(original.Start, deserialized.Start);
    }

    [Fact]
    public void AttachmentSaveResult_WithSavedFiles_RoundTrips()
    {
        var original = new AttachmentSaveResult
        {
            Success = true,
            EntryId = "entry-2",
            StoreId = "store-2",
            Subject = "Attachment mail",
            SavedCount = 2,
            SavedFiles = [@"C:\temp\a.txt", @"C:\temp\b.txt"],
            Message = "Saved 2 Outlook attachments."
        };

        var json = JsonSerializer.Serialize(original, JsonOptions);
        var deserialized = JsonSerializer.Deserialize<AttachmentSaveResult>(json, JsonOptions);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Success, deserialized.Success);
        Assert.Equal(original.EntryId, deserialized.EntryId);
        Assert.Equal(original.StoreId, deserialized.StoreId);
        Assert.Equal(original.Subject, deserialized.Subject);
        Assert.Equal(original.SavedCount, deserialized.SavedCount);
        Assert.Equal(original.SavedFiles, deserialized.SavedFiles);
        Assert.Equal(original.Message, deserialized.Message);
    }

    [Fact]
    public void AttachmentMutationResult_WithFileName_RoundTrips()
    {
        var original = new AttachmentMutationResult
        {
            Success = true,
            EntryId = "entry-3",
            StoreId = "store-3",
            Subject = "Draft with attachment",
            AttachmentCount = 1,
            FileName = "demo.txt",
            Message = "Added Outlook attachment to draft."
        };

        var json = JsonSerializer.Serialize(original, JsonOptions);
        var deserialized = JsonSerializer.Deserialize<AttachmentMutationResult>(json, JsonOptions);

        Assert.NotNull(deserialized);
        Assert.Equal(original.EntryId, deserialized.EntryId);
        Assert.Equal(original.StoreId, deserialized.StoreId);
        Assert.Equal(original.Subject, deserialized.Subject);
        Assert.Equal(original.AttachmentCount, deserialized.AttachmentCount);
        Assert.Equal(original.FileName, deserialized.FileName);
        Assert.Equal(original.Message, deserialized.Message);
    }

    [Fact]
    public void ActiveMailResult_WithoutOptionalFields_OmitsNullProperties()
    {
        var result = new ActiveMailResult
        {
            Success = true,
            HasActiveMail = false,
            Unread = false,
            Importance = 0,
            AttachmentCount = 0
        };

        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"hasActiveMail\":false", json);
        Assert.Contains("\"attachmentCount\":0", json);
        Assert.DoesNotContain("\"entryId\":", json);
        Assert.DoesNotContain("\"bodyPreview\":", json);
        Assert.DoesNotContain("\"receivedTime\":", json);
    }
}
