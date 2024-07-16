using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Reporting.NETCore;
using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;

LocalReport? report = new LocalReport();
//Load(report);
//Load2(report);
//byte[]? pdf = report.Render("PDF");
//byte[]? word = report.Render("WORDOPENXML");

List<byte[]> listBytes = new List<byte[]>();
for (int i = 1; i <= 3; i++)
{
    Load3(report, i);
    byte[]? word = report.Render("WORDOPENXmL");
    listBytes.Add(word!);

    report.Dispose();
    report = new LocalReport();
}

//int contador = 1;
//foreach(var item in listBytes)
//{
//    File.WriteAllBytes($"{contador}_{DateTime.Now.ToString("yyyyMMdd-HHmmss")}.docx", item);
//    contador++;
//}

//byte[]? result = null;
//for(int i = 0; i < listBytes.Count; i++)
//{
//    if (i == 0)
//    {
//        result = listBytes[i];
//    }

//    result = MergeDo
//}

byte[] result = MergeDocuments(listBytes);


//File.WriteAllBytes("report.pdf", pdf);
//File.WriteAllBytes($"report_{DateTime.Now.ToString("yyyyMMdd-HHmmss")}.docx", word);
File.WriteAllBytes($"reportMerged_{DateTime.Now.ToString("yyyyMMdd-HHmmss")}.docx", result);

void Load(LocalReport report)
{
    ReportItem[] items = new[] { new ReportItem { Description = "Widget 6000", Price = 104.99m, Qty = 1 }, new ReportItem { Description = "Gizmo MAX", Price = 1.41m, Qty = 25 } };
    ReportParameter[] parameters = new[] { new ReportParameter("Title", "Invoice 4/2020") };
    using (FileStream fs = new FileStream("Reports/Console-Sample-Test.rdlc", FileMode.Open))
    {
        report.LoadReportDefinition(fs);
        report.DataSources.Add(new ReportDataSource("Items", items));
        report.SetParameters(parameters);
    }
}

void Load2(LocalReport report)
{
    byte[] imagen1 = File.ReadAllBytes(@"Imgs/1.png");
    byte[] imagen2 = File.ReadAllBytes(@"Imgs/2.png");
    byte[] imagen3 = File.ReadAllBytes(@"Imgs/3.png");

    byte[] imagen4 = File.ReadAllBytes(@"Imgs/4.png");
    byte[] imagen5 = File.ReadAllBytes(@"Imgs/5.png");
    byte[] imagen6 = File.ReadAllBytes(@"Imgs/6.png");

    byte[] imagen7 = File.ReadAllBytes(@"Imgs/7.png");
    byte[] imagen8 = File.ReadAllBytes(@"Imgs/8.png");
    byte[] imagen9 = File.ReadAllBytes(@"Imgs/9.png");

    FigurantesPresentacion[] items = new[] {
        new FigurantesPresentacion
        {
            Imagen1 = Convert.ToBase64String(imagen1),
            Imagen2 = Convert.ToBase64String(imagen2),
            Imagen3 = Convert.ToBase64String(imagen3),
            NumeroCasting = "OM00110",
            Edad = 25,
            Altura = 184,
            Chaqueta = "XL",
            Pantalon = 46,
            Zapato = "45",
            NombreCompleto = "Daniel Santos Mundiña",
            Personaje = "Jugador de rugby"
        }
        ,
        new FigurantesPresentacion
        {
            Imagen1 = Convert.ToBase64String(imagen4),
            Imagen2 = Convert.ToBase64String(imagen5),
            Imagen3 = Convert.ToBase64String(imagen6),
            NumeroCasting = "OM00111",
            Edad = 27,
            Altura = 179,
            Chaqueta = "XL",
            Pantalon = 42,
            Zapato = "44",
            NombreCompleto = "Juan Fernández Nuñez",
            Personaje = "Jugador de rugby"
        }
        ,
        new FigurantesPresentacion
        {
            Imagen1 = Convert.ToBase64String(imagen7),
            Imagen2 = Convert.ToBase64String(imagen8),
            Imagen3 = string.Empty,
            NumeroCasting = "OM00118",
            Edad = 32,
            Altura = 174,
            Chaqueta = "L",
            Pantalon = 40,
            Zapato = "43",
            NombreCompleto = "Samuel Rodríguez Perez",
            Personaje = "Jugador de rugby"
        }
    };
    ReportParameter[] parameters = new[] { new ReportParameter("Title", "OLYMPO T1") };
    using (FileStream fs = new FileStream("Reports/Test-Tabla.rdlc", FileMode.Open))
    {
        report.LoadReportDefinition(fs);
        report.DataSources.Add(new ReportDataSource("Items", items));
        report.SetParameters(parameters);
    }
}

void Load3(LocalReport report, int alternor)
{
    byte[] imagen1 = File.ReadAllBytes(@"Imgs/1.png");
    byte[] imagen2 = File.ReadAllBytes(@"Imgs/2.png");
    byte[] imagen3 = File.ReadAllBytes(@"Imgs/3.png");

    byte[] imagen4 = File.ReadAllBytes(@"Imgs/4.png");
    byte[] imagen5 = File.ReadAllBytes(@"Imgs/5.png");
    byte[] imagen6 = File.ReadAllBytes(@"Imgs/6.png");

    byte[] imagen7 = File.ReadAllBytes(@"Imgs/7.png");
    byte[] imagen8 = File.ReadAllBytes(@"Imgs/8.png");
    byte[] imagen9 = File.ReadAllBytes(@"Imgs/9.png");

    List<FigurantesPresentacion> items = new List<FigurantesPresentacion>();

    if (alternor == 1)
    {
        items.Add(new FigurantesPresentacion
        {
            Imagen1 = Convert.ToBase64String(imagen1),
            Imagen2 = Convert.ToBase64String(imagen2),
            Imagen3 = Convert.ToBase64String(imagen3),
            NumeroCasting = "OM00110",
            Edad = 25,
            Altura = 184,
            Chaqueta = "XL",
            Pantalon = 46,
            Zapato = "45",
            NombreCompleto = "Daniel Santos Mundiña",
            Personaje = "Jugador de rugby"
        });
    }

    if(alternor == 2)
    {
        items.Add(new FigurantesPresentacion
        {
            Imagen1 = Convert.ToBase64String(imagen4),
            Imagen2 = Convert.ToBase64String(imagen5),
            Imagen3 = Convert.ToBase64String(imagen6),
            NumeroCasting = "OM00111",
            Edad = 27,
            Altura = 179,
            Chaqueta = "XL",
            Pantalon = 42,
            Zapato = "44",
            NombreCompleto = "Juan Fernández Nuñez",
            Personaje = "Jugador de rugby"
        });
    }

    if(alternor == 3)
    {
        items.Add(new FigurantesPresentacion
        {
            Imagen1 = Convert.ToBase64String(imagen7),
            Imagen2 = Convert.ToBase64String(imagen8),
            Imagen3 = string.Empty,
            NumeroCasting = "OM00118",
            Edad = 32,
            Altura = 174,
            Chaqueta = "L",
            Pantalon = 40,
            Zapato = "43",
            NombreCompleto = "Samuel Rodríguez Perez",
            Personaje = "Jugador de rugby"
        });
    }

    ReportParameter[] parameters = new[] { new ReportParameter("Title", "OLYMPO T1") };
    using (FileStream fs = new FileStream("Reports/Test-Tabla.rdlc", FileMode.Open))
    {
        report.LoadReportDefinition(fs);
        report.DataSources.Add(new ReportDataSource("Items", items));
        report.SetParameters(parameters);
    }
}

static byte[] MergeDocuments(List<byte[]> docxFiles)
{
    using (MemoryStream outputStream = new MemoryStream())
    {
        using (WordprocessingDocument outputDoc = WordprocessingDocument.Create(outputStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Añadir la parte principal del documento
            MainDocumentPart mainPart = outputDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body outputBody = mainPart.Document.AppendChild(new Body());

            foreach (byte[] docxFile in docxFiles)
            {
                using (MemoryStream inputStream = new MemoryStream(docxFile))
                using (WordprocessingDocument inputDoc = WordprocessingDocument.Open(inputStream, false))
                {
                    // Clonar elementos del cuerpo del documento
                    Body inputBody = inputDoc.MainDocumentPart.Document.Body.CloneNode(true) as Body;
                    foreach (var element in inputBody.Elements())
                    {
                        outputBody.AppendChild(element.CloneNode(true));
                    }

                    // Copiar las partes relacionadas, como imágenes
                    CopyRelatedParts(inputDoc.MainDocumentPart, mainPart);
                }
            }

            mainPart.Document.Save();
        }

        return outputStream.ToArray();
    }
}

static void CopyRelatedParts(MainDocumentPart sourcePart, MainDocumentPart targetPart)
{
    foreach (var part in sourcePart.Parts)
    {
        if (part.OpenXmlPart is ImagePart)
        {
            var imagePart = (ImagePart)part.OpenXmlPart;
            var newImagePart = targetPart.AddImagePart(imagePart.ContentType);

            using (var stream = imagePart.GetStream())
            {
                newImagePart.FeedData(stream);
            }

            UpdateImageReferences(sourcePart, targetPart, part.RelationshipId, targetPart.GetIdOfPart(newImagePart));
        }
        else
        {
            if (!targetPart.Parts.Any(p => p.OpenXmlPart.Uri == part.OpenXmlPart.Uri))
            {
                targetPart.AddPart(part.OpenXmlPart);
            }
        }
    }
}

static void UpdateImageReferences(MainDocumentPart sourcePart, MainDocumentPart targetPart, string oldRelId, string newRelId)
{
    foreach (var drawing in targetPart.Document.Body.Descendants<Drawing>())
    {
        var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        if (blip != null && blip.Embed == oldRelId)
        {
            blip.Embed = newRelId;
        }
    }
}

//byte[] MergeDocuments(List<byte[]> inputFiles)
//{
//    using (MemoryStream outputStream = new MemoryStream())
//    {
//        using (WordprocessingDocument outputDoc = WordprocessingDocument.Create(outputStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
//        {
//            // Añadir la parte principal del documento
//            MainDocumentPart mainPart = outputDoc.AddMainDocumentPart();
//            mainPart.Document = new Document();
//            Body outputBody = mainPart.Document.AppendChild(new Body());

//            int contador = 1000;
//            foreach (byte[] docxFile in inputFiles)
//            {
//                using (MemoryStream inputStream = new MemoryStream(docxFile))
//                using (WordprocessingDocument inputDoc = WordprocessingDocument.Open(inputStream, false))
//                {
//                    // Clonar elementos del cuerpo del documento
//                    Body inputBody = inputDoc.MainDocumentPart.Document.Body.CloneNode(true) as Body;
//                    foreach (var element in inputBody.Elements())
//                    {
//                        outputBody.AppendChild(element.CloneNode(true));
//                    }

//                    // Copiar las partes relacionadas, como imágenes
//                    foreach (var part in inputDoc.MainDocumentPart.Parts)
//                    {
//                        contador++;
//                        string relId = $"{part.RelationshipId}_{contador}";
//                        if (!mainPart.Parts.Any(p => p.RelationshipId == part.RelationshipId && part.RelationshipId == relId))
//                        {
//                            var partType = part.OpenXmlPart.ContentType;
//                            try
//                            {
//                                var targetPart = mainPart.AddPart(part.OpenXmlPart, relId);

//                                contador++;
//                                relId = $"{part.RelationshipId}_{contador}";
//                                mainPart.ChangeIdOfPart(targetPart, relId);
//                            }
//                            catch
//                            {
//                                //contador++;
//                                //relId = $"{part.RelationshipId}_{contador}";
//                                //var targetPart = mainPart.AddPart(part.OpenXmlPart, relId);
//                                //mainPart.ChangeIdOfPart(targetPart, relId);
//                            }
//                        }
//                        contador++;
//                    }
//                }
//                contador++;
//            }

//            mainPart.Document.Save();
//        }

//        return outputStream.ToArray();
//    }
//}

class ReportItem
{
    public string Description { get; set; }
    public decimal Price { get; set; }
    public int Qty { get; set; }
    public decimal Total => Price * Qty;
}

class FigurantesPresentacion
{
    public string? Imagen1 { get; set; }
    public string? Imagen2 { get; set; }
    public string? Imagen3 { get; set; }
    public string NumeroCasting { get; set; }
    public int Edad { get; set; }
    public int Altura { get; set; }
    public string Chaqueta { get; set; }
    public int Pantalon { get; set; }
    public string Zapato { get; set; }
    public string NombreCompleto { get; set; }
    public string Personaje { get; set; }
}