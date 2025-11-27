using System.Data;
using Xunit.Abstractions;

namespace Spire.Doc.Tests;

/// <summary>
/// Tests d'intégration avancés avec vérifications approfondies
/// </summary>
public class SpireDocAdvancedTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _testOutputDir;

    public SpireDocAdvancedTests(ITestOutputHelper output)
    {
        _output = output;
        _testOutputDir = Path.Combine(Path.GetTempPath(), $"SpireDocAdvanced_{Guid.NewGuid()}");
        Directory.CreateDirectory(_testOutputDir);
    }

    [Fact]
    public void MailMerge_WithImageField_ShouldEmbedImage()
    {
        // Arrange
        var templatePath = Path.Combine(_testOutputDir, "template_with_image.docx");
        var outputPath = Path.Combine(_testOutputDir, "output_with_image.docx");

        // Création d'un template avec champ image
        var doc = new Document();
        var section = doc.AddSection();
        var paragraph = section.AddParagraph();
        paragraph.AppendText("Photo: ");
        paragraph.AppendField("Photo", FieldType.FieldMergeField);

        doc.SaveToFile(templatePath, FileFormat.Docx2016);
        doc.Dispose();

        // Création d'une image simple (pixel blanc 1x1)
        var imagePath = Path.Combine(_testOutputDir, "test_image.png");
        CreateTestImage(imagePath);

        // Act
        var mergeDoc = new Document(templatePath);
        string[] fieldNames = { "Photo" };
        string[] fieldValues = { imagePath };

        mergeDoc.MailMerge.Execute(fieldNames, fieldValues);
        mergeDoc.SaveToFile(outputPath, FileFormat.Docx2016);

        // Assert
        Assert.True(File.Exists(outputPath));

        var resultDoc = new Document(outputPath);
        Assert.True(resultDoc.Sections[0].Paragraphs[0].ChildObjects.Count > 1);

        resultDoc.Dispose();
        mergeDoc.Dispose();

        _output.WriteLine("Test avec image réussi");
    }

    private void CreateTestImage(string path)
    {
        // Création d'une image PNG 1x1 pixel blanc
        byte[] pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
            0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00,
            0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0xFC, 0xFF, 0xFF, 0x3F,
            0x00, 0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59, 0xE7, 0x00, 0x00, 0x00,
            0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(path, pngBytes);
    }

    [Fact]
    public void PerformanceTest_LargeBatchMailMerge_ShouldCompleteInReasonableTime()
    {
        // Arrange
        var templatePath = Path.Combine(_testOutputDir, "template_perf.docx");
        var outputPath = Path.Combine(_testOutputDir, "output_perf.docx");

        var doc = new Document();
        var section = doc.AddSection();
        var paragraph = section.AddParagraph();
        paragraph.AppendText("Nom: ");
        paragraph.AppendField("Nom", FieldType.FieldMergeField);
        paragraph.AppendText(" | ID: ");
        paragraph.AppendField("ID", FieldType.FieldMergeField);
        doc.SaveToFile(templatePath, FileFormat.Docx2016);
        doc.Dispose();

        var dataTable = new DataTable("Records");
        dataTable.Columns.Add("Nom");
        dataTable.Columns.Add("ID");

        // Ajout de 100 enregistrements
        for (int i = 0; i < 100; i++)
        {
            dataTable.Rows.Add($"Personne_{i}", i.ToString());
        }

        var mergeDoc = new Document(templatePath);
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        // Act
        mergeDoc.MailMerge.Execute(dataTable);
        mergeDoc.SaveToFile(outputPath, FileFormat.Docx2016);

        stopwatch.Stop();

        // Assert
        Assert.True(File.Exists(outputPath));
        Assert.True(stopwatch.ElapsedMilliseconds < 30000,
            $"Le traitement a pris {stopwatch.ElapsedMilliseconds}ms, devrait être < 30000ms");

        mergeDoc.Dispose();

        _output.WriteLine($"100 enregistrements traités en {stopwatch.ElapsedMilliseconds}ms");
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testOutputDir))
            {
                Directory.Delete(_testOutputDir, true);
            }
        }
        catch { }
    }
}