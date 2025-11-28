using System.Data;
using Xunit.Abstractions;

namespace Spire.Doc.Tests;

/// <summary>
/// Tests d'intégration pour le publipostage avec Spire.Doc
/// Package requis: dotnet add package Spire.Doc
/// </summary>
public class SpireDocMailMergeTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _testOutputDir;
    private readonly string _templatePath;

    public SpireDocMailMergeTests(ITestOutputHelper output)
    {
        _output = output;
        _testOutputDir = Path.Combine(Path.GetTempPath(), $"SpireDocTests_{Guid.NewGuid()}");
        Directory.CreateDirectory(_testOutputDir);

        _templatePath = Path.Combine(_testOutputDir, "template.docx");
        CreateMailMergeTemplate();
    }

    /// <summary>
    /// Crée un template Word avec des champs de fusion
    /// </summary>
    private void CreateMailMergeTemplate()
    {
        var doc = new Document();
        var section = doc.AddSection();
        var paragraph = section.AddParagraph();

        // Ajout de texte et champs de fusion
        paragraph.AppendText("Lettre de bienvenue\n\n");
        paragraph.AppendText("Cher(e) ");
        paragraph.AppendField("Prenom", FieldType.FieldMergeField);
        paragraph.AppendText(" ");
        paragraph.AppendField("Nom", FieldType.FieldMergeField);
        paragraph.AppendText(",\n\n");

        paragraph.AppendText("Nous sommes heureux de vous accueillir dans notre entreprise ");
        paragraph.AppendField("Entreprise", FieldType.FieldMergeField);
        paragraph.AppendText(".\n\n");

        paragraph.AppendText("Votre poste: ");
        paragraph.AppendField("Poste", FieldType.FieldMergeField);
        paragraph.AppendText("\n");

        paragraph.AppendText("Date de début: ");
        paragraph.AppendField("DateDebut", FieldType.FieldMergeField);
        paragraph.AppendText("\n\n");

        paragraph.AppendText("Cordialement,\nLes Ressources Humaines");

        doc.SaveToFile(_templatePath, FileFormat.Docx2016);
        doc.Dispose();

        _output.WriteLine($"Template créé: {_templatePath}");
    }

    [Fact]
    public void MailMerge_WithSingleRecord_ShouldGenerateDocument()
    {
        // Arrange
        var outputPath = Path.Combine(_testOutputDir, "output_single.docx");
        var doc = new Document(_templatePath);

        string[] fieldNames = { "Prenom", "Nom", "Entreprise", "Poste", "DateDebut" };
        string[] fieldValues = { "Anthony", "Coudène", "TechCorp", "Développeur Senior", "01/01/2025" };

        // Act
        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.SaveToFile(outputPath, FileFormat.Docx2016);

        // Assert
        Assert.True(File.Exists(outputPath));

        // Vérification du contenu
        var resultDoc = new Document(outputPath);
        string text = resultDoc.GetText();

        Assert.Contains("Anthony Coudène", text);
        Assert.Contains("TechCorp", text);
        Assert.Contains("Développeur Senior", text);
        Assert.Contains("01/01/2025", text);
        Assert.DoesNotContain("«Prenom»", text); // Les champs doivent être remplacés

        resultDoc.Dispose();
        doc.Dispose();

        _output.WriteLine($"Document généré: {outputPath}");
        _output.WriteLine($"Contenu vérifié avec succès");
    }

    [Fact]
    public void MailMerge_WithDataTable_ShouldGenerateMultipleDocuments()
    {
        // Arrange
        var dataTable = new DataTable("Employees");
        dataTable.Columns.Add("Prenom");
        dataTable.Columns.Add("Nom");
        dataTable.Columns.Add("Entreprise");
        dataTable.Columns.Add("Poste");
        dataTable.Columns.Add("DateDebut");

        dataTable.Rows.Add("Marie", "Martin", "DevCorp", "Chef de Projet", "15/01/2025");
        dataTable.Rows.Add("Anthony", "Coudène", "CodeInc", "Architecte", "01/02/2025");
        dataTable.Rows.Add("Sophie", "Bernard", "WebSoft", "Designer UX", "10/02/2025");

        var outputPath = Path.Combine(_testOutputDir, "output_multiple.docx");
        var doc = new Document(_templatePath);

        // Act
        doc.MailMerge.Execute(dataTable);
        doc.SaveToFile(outputPath, FileFormat.Docx2016);

        // Assert
        Assert.True(File.Exists(outputPath));

        var resultDoc = new Document(outputPath);
        string text = resultDoc.GetText();

        // Vérification que tous les enregistrements sont présents
        Assert.Contains("Marie Martin", text);
        Assert.Contains("Anthony Coudène", text);
        Assert.Contains("Sophie Bernard", text);
        Assert.Contains("DevCorp", text);
        Assert.Contains("CodeInc", text);
        Assert.Contains("WebSoft", text);

        resultDoc.Dispose();
        doc.Dispose();

        _output.WriteLine($"Document multiple généré: {outputPath}");
        _output.WriteLine($"3 enregistrements fusionnés avec succès");
    }

    [Fact]
    public void MailMerge_AndConvertToPDF_ShouldGeneratePDFFile()
    {
        // Arrange
        var outputDocx = Path.Combine(_testOutputDir, "output_for_pdf.docx");
        var outputPdf = Path.Combine(_testOutputDir, "output.pdf");
        var doc = new Document(_templatePath);

        string[] fieldNames = { "Prenom", "Nom", "Entreprise", "Poste", "DateDebut" };
        string[] fieldValues = { "Luc", "Petit", "InnoTech", "Data Scientist", "20/01/2025" };

        // Act
        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.SaveToFile(outputDocx, FileFormat.Docx2016);

        // Conversion en PDF
        doc.SaveToFile(outputPdf, FileFormat.PDF);

        // Assert
        Assert.True(File.Exists(outputPdf));

        var pdfInfo = new FileInfo(outputPdf);
        Assert.True(pdfInfo.Length > 0, "Le fichier PDF doit avoir une taille supérieure à 0");

        doc.Dispose();

        _output.WriteLine($"PDF généré: {outputPdf}");
        _output.WriteLine($"Taille du PDF: {pdfInfo.Length} bytes");
    }

    [Fact]
    public void MailMerge_AndConvertToText_ShouldGenerateTextFile()
    {
        // Arrange
        var outputTxt = Path.Combine(_testOutputDir, "output.txt");
        var doc = new Document(_templatePath);

        string[] fieldNames = { "Prenom", "Nom", "Entreprise", "Poste", "DateDebut" };
        string[] fieldValues = { "Claire", "Moreau", "CloudSystem", "DevOps Engineer", "05/02/2025" };

        // Act
        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.SaveToFile(outputTxt, FileFormat.Txt);

        // Assert
        Assert.True(File.Exists(outputTxt));

        string content = File.ReadAllText(outputTxt);
        Assert.Contains("Claire Moreau", content);
        Assert.Contains("CloudSystem", content);
        Assert.Contains("DevOps Engineer", content);

        doc.Dispose();

        _output.WriteLine($"Fichier texte généré: {outputTxt}");
        _output.WriteLine($"Contenu:\n{content}");
    }

    [Fact]
    public void MailMerge_WithMissingFields_ShouldHandleGracefully()
    {
        // Arrange
        var outputPath = Path.Combine(_testOutputDir, "output_missing_fields.docx");
        var doc = new Document(_templatePath);

        // Seulement 2 champs sur 5
        string[] fieldNames = { "Prenom", "Nom" };
        string[] fieldValues = { "Anthony", "Coudène" };

        // Act
        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.SaveToFile(outputPath, FileFormat.Docx2016);

        // Assert
        Assert.True(File.Exists(outputPath));

        var resultDoc = new Document(outputPath);
        string text = resultDoc.GetText();

        // Les champs fournis doivent être remplis
        Assert.Contains("Anthony Coudène", text);

        // Les champs non fournis restent invisibles
        Assert.DoesNotContain("«Entreprise»", text);
        Assert.DoesNotContain("«Poste»", text);

        resultDoc.Dispose();
        doc.Dispose();

        _output.WriteLine($"Test des champs manquants réussi");
    }

    [Fact]
    public void MailMerge_WithComplexDataTypes_ShouldFormatCorrectly()
    {
        // Arrange
        var outputPath = Path.Combine(_testOutputDir, "output_complex.docx");
        var doc = new Document(_templatePath);

        var dateDebut = new DateTime(2025, 3, 15);
        string[] fieldNames = { "Prenom", "Nom", "Entreprise", "Poste", "DateDebut" };
        string[] fieldValues =
        {
            "Émilie",
            "Lefèvre",
            "Société & Co.",
            "Expert·e Technique",
            dateDebut.ToString("dd/MM/yyyy")
        };

        // Act
        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.SaveToFile(outputPath, FileFormat.Docx2016);

        // Assert
        Assert.True(File.Exists(outputPath));

        var resultDoc = new Document(outputPath);
        string text = resultDoc.GetText();

        // Vérification des caractères spéciaux et formatage
        Assert.Contains("Émilie Lefèvre", text);
        Assert.Contains("Société & Co.", text);
        Assert.Contains("Expert·e Technique", text);
        Assert.Contains("15/03/2025", text);

        resultDoc.Dispose();
        doc.Dispose();

        _output.WriteLine("Gestion des caractères spéciaux réussie");
    }

    public void Dispose()
    {
        // Nettoyage des fichiers de test
        try
        {
            if (Directory.Exists(_testOutputDir))
            {
                Directory.Delete(_testOutputDir, true);
                _output.WriteLine($"Répertoire de test nettoyé: {_testOutputDir}");
            }
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Erreur lors du nettoyage: {ex.Message}");
        }
    }
}

