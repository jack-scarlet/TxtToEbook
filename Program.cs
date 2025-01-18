using HtmlAgilityPack;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.IO.Compression;
using System.Reflection.Metadata;
using System.Text;
using System.Web;
using System.Xml;

namespace TextToEpub
{
    class Program
    {
        static string directory = @"C:\Caminho\Para\Os\Arquivos";
        static string outputFileName = "output.epub";
        static string tempDirectory = "temp";
        static string BookName = "";
        static string BookAuthor = "";
        static string BookGUID = Guid.NewGuid().ToString();
        static void Main(string[] args)
        {

            if (args.Length != 1)
                return;

            directory = args.First();

            outputFileName = args.Last().TrimEnd(' ', '/', '\\') + ".epub";


            Console.WriteLine("Type the ebook name:");
            BookName = Console.ReadLine() ?? throw new Exception();
            Console.WriteLine("Type the Author name:");
            BookAuthor = Console.ReadLine() ?? throw new Exception();

            if (Directory.GetFiles(directory, "*.btxt").Length > 0)
            {
                var Files = Directory.GetFiles(directory, "*.btxt");
                foreach (var Book in Files)
                {
                    var Lines = File.ReadAllLines(Book);

                    var Trimmer = (string x) => x.Trim().Trim('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');

                    var DummyLine = (string x) => x.Length > 0 && x.Trim(x.Trim().First(), ' ', '\t') == string.Empty;

                    var ChapterPrefix = Lines.First(x => !string.IsNullOrWhiteSpace(x));
                    ChapterPrefix = Trimmer(ChapterPrefix);

                    string imageFile = null, imageFileAlt = null;

                    List<string> ChapterContent = new List<string>();
                    for (int i = 0, chp = 0, breaks = 0; i < Lines.Length; i++)
                    {
                        var Line = Lines[i];

                        var NextLine = i + 1 < Lines.Length ? Lines[i + 1] : null;

                        var Prefix = Trimmer(Line);

                        if (Prefix == ChapterPrefix)
                        {
                            if (ChapterContent.Count > 0)
                            {
                                if (imageFile != null)
                                {
                                    ChapterContent.Add($"<p><img src=\"img/{Path.GetFileName(imageFile)}\" alt=\"Chapter {chp} Illustration\"/></p>");
                                    imageFile = null;
                                }

                                if (imageFileAlt != null)
                                {
                                    ChapterContent.Add($"<p><img src=\"img/{Path.GetFileName(imageFileAlt)}\" alt=\"Chapter {chp} Illustration\"/></p>");
                                    imageFileAlt = null;
                                }

                                string SaveAs = Path.Combine(directory, $"{chp}.txt");
                                File.WriteAllLines(SaveAs, ChapterContent.ToArray());
                                ChapterContent.Clear();
                            }

                            chp = int.Parse(Line.Replace(Prefix, "").Trim());

                            // Verificar se há uma imagem correspondente
                            imageFile = Path.Combine(directory, "img", $"{chp}.jpg");
                            if (!File.Exists(imageFile))
                            {
                                imageFile = null;
                            }

                            // Verificar se há uma imagem adicional correspondente
                            imageFileAlt = Path.Combine(directory, "img", $"{chp}-2.jpg");
                            if (!File.Exists(imageFileAlt))
                            {
                                imageFileAlt = null;
                            }

                            continue;
                        }

                        if (Line == "[image]" && imageFile != null)
                        {
                            ChapterContent.Add($"<p><img src=\"img/{Path.GetFileName(imageFile)}\" alt=\"Chapter {chp} Illustration\"/></p>");
                            imageFile = null;
                            continue;
                        }

                        if (Line == "[image]" && imageFileAlt != null)
                        {
                            ChapterContent.Add($"<p><img src=\"img/{Path.GetFileName(imageFileAlt)}\" alt=\"Chapter {chp} Illustration\"/></p>");
                            imageFileAlt = null;
                            continue;
                        }

                        if (!string.IsNullOrWhiteSpace(Line))
                        {
                            if (Line.StartsWith("**"))
                                Line = " " + Line;

                            if (Line.EndsWith("**"))
                                Line += " ";

                            ChapterContent.Add($"<p>{HttpUtility.HtmlEncode(Line).Replace(" **", "<i>").Replace("** ", "</i>")}</p>");

                            if (DummyLine(Line))
                                breaks = -3;
                            else
                                breaks = 0;
                        }
                        else
                        {
                            if (ChapterContent.Count > 0)
                            {
                                if (ChapterContent.Count == 0)
                                    ChapterContent.Add("<p><br /></p>");
                                else
                                    ChapterContent[ChapterContent.Count - 1] = ChapterContent.Last().Replace("</p>", "<br /></p>");

                                breaks++;
                            }
                            else
                                continue;
                        }

                        if (breaks == 3)
                        {
                            if (DummyLine(NextLine) || NextLine.StartsWith('['))
                            {
                                breaks = 0;
                                continue;
                            }
                        }
                    }

                }
            }

            // Criar diretório temporário
            Directory.CreateDirectory(tempDirectory);

            // Separar conteúdo em capítulos
            var chapters = Directory.GetFiles(directory, "*.txt").OrderBy(x => int.Parse(Path.GetFileNameWithoutExtension(x))).ToArray();

            // Processar cada arquivo TXT
            for (int i = 0; i < chapters.Length; i++) 
            {
                var file = chapters[i];

                // Ler conteúdo do arquivo TXT
                string content = File.ReadAllText(file);

                // Criar arquivo HTML para o capítulo
                string chapterHtmlPath = Path.Combine(tempDirectory, $"chapter_{i + 1}.html");
                File.WriteAllText(chapterHtmlPath, $"<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Chapter {i + 1}</title></head><body><h1>Chapter {i + 1}</h1>\n{File.ReadAllText(chapters[i])}</body></html>");
            }

            if (File.Exists(outputFileName))
                File.Delete(outputFileName);

            // Criar arquivo EPUB
            using (ZipArchive epub = ZipFile.Open(outputFileName, ZipArchiveMode.Create))
            {
                // Adicionar mime
                var mimeEntry = epub.CreateEntry("mimetype", CompressionLevel.NoCompression);

                using (var Stream = new StreamWriter(mimeEntry.Open()))
                {
                    Stream.Write("application/epub+zip");
                }

                // Adicionar metadados
                AddMetadata(epub);

                // Adicionar conteúdo
                foreach (string file in Directory.GetFiles(tempDirectory, "*.html").OrderBy(x => int.Parse(Path.GetFileNameWithoutExtension(x).Replace("chapter_", ""))))
                {
                    epub.CreateEntryFromFile(file, $"OEBPS/{Path.GetFileName(file)}");
                }

                // Adicionar conteúdo
                foreach (string file in Directory.GetFiles(directory, "*.jpg", SearchOption.AllDirectories))
                {
                    epub.CreateEntryFromFile(file, $"OEBPS/img/{Path.GetFileName(file)}");
                }
            }

            // Excluir diretório temporário
            Directory.Delete(tempDirectory, true);

            Console.WriteLine($"EPUB criado com sucesso: {outputFileName}");
            Console.ReadLine();
        }

        static void AddMetadata(ZipArchive epub)
        {
            // Criar arquivo container.xml
            string containerXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>";
            ZipArchiveEntry containerEntry = epub.CreateEntry("META-INF/container.xml");
            using (StreamWriter writer = new StreamWriter(containerEntry.Open()))
            {
                writer.Write(containerXml);
            }

            string coverPage = "<?xml version='1.0' encoding='utf-8'?><html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:lang=\"en\"><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"/><meta name=\"calibre:cover\" content=\"true\"/><title>Cover</title><style type=\"text/css\" title=\"override_css\">@page {padding: 0pt; margin:0pt} body { text-align: center; padding:0pt; margin: 0pt; }</style></head><body><div><img width=\"1200\" height=\"1600\" src=\"img/cover.jpg\" alt=\"Book Cover\"/></div></body></html>";
            ZipArchiveEntry coverEntry = epub.CreateEntry("OEBPS/cover.xhtml");
            using (StreamWriter writer = new StreamWriter(coverEntry.Open()))
            {
                writer.Write(coverPage);
            }

            // Criar arquivo content.opf
            var contentOpf = new HtmlDocument();
            contentOpf.OptionOutputAsXml = true;
            contentOpf.OptionPreserveXmlNamespaces = true;
            contentOpf.OptionAutoCloseOnEnd = false;
            contentOpf.OptionCheckSyntax = false;
            contentOpf.OptionXmlForceOriginalComment = true;
            contentOpf.OptionOutputOriginalCase = true;
            contentOpf.LoadHtml($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><package xmlns=\"http://www.idpf.org/2007/opf\" version=\"2.0\" unique-identifier=\"uuid_id\"><metadata><dc:title>{BookName}</dc:title><dc:creator>{BookAuthor}</dc:creator><dc:language>en</dc:language><dc:publisher>{BookAuthor}</dc:publisher><dc:rights>Creative Commons</dc:rights></metadata><manifest></manifest><spine toc=\"ncx\"></spine></package>");

            CreateNamespace(contentOpf);
            SetOPFMetadata(contentOpf);

            var manifestNode = contentOpf.DocumentNode.SelectSingleNode("//manifest");
            var spineNode = contentOpf.DocumentNode.SelectSingleNode("//spine");


            foreach (string file in Directory.GetFiles(tempDirectory, "*.html").OrderBy(x => int.Parse(Path.GetFileNameWithoutExtension(x).Replace("chapter_", ""))))
            {
                string fileName = Path.GetFileName(file);
                manifestNode.AppendChild(CreateManifestItemNode(contentOpf, fileName));
                spineNode.AppendChild(CreateSpineItemNode(contentOpf, fileName));
            }

            foreach (var image in Directory.GetFiles(directory, "*.jpg", SearchOption.AllDirectories))
            {
                manifestNode.AppendChild(CreateManifestItemNode(contentOpf, $"img/{Path.GetFileName(image)}", "image/jpeg", $"img_{Path.GetFileNameWithoutExtension(image)}"));
            }

            if (File.Exists(Path.Combine(directory, "cover.jpg")))
            {
                manifestNode.AppendChild(CreateManifestItemNode(contentOpf, "cover.xhtml"));
                spineNode.InsertBefore(CreateSpineItemNode(contentOpf, "cover.xhtml"), spineNode.FirstChild);
            }

            manifestNode.AppendChild(CreateManifestItemNode(contentOpf, "toc.ncx", "application/x-dtbncx+xml", "ncx"));


            ZipArchiveEntry contentOpfEntry = epub.CreateEntry("OEBPS/content.opf");

            using (MemoryStream buffer = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(buffer))
            using (StreamReader reader = new StreamReader(buffer))
            using (StreamWriter Outwriter = new StreamWriter(contentOpfEntry.Open()))
            {
                contentOpf.Save(writer);
                buffer.Flush();
                buffer.Position = 0;

                var XML = reader.ReadToEnd();
                XML = XML.Replace("<?xml version=\"1.0\" encoding=\"utf-8\"?><span>", "");
                XML = XML.Replace("</span>", "");

                Outwriter.Write(XML);
            }

            // Criar arquivo toc.ncx
            var tocNcx = new HtmlDocument();
            tocNcx.OptionOutputAsXml = true;
            tocNcx.OptionPreserveXmlNamespaces = true;
            tocNcx.OptionAutoCloseOnEnd = false;
            tocNcx.OptionCheckSyntax = false;
            tocNcx.OptionXmlForceOriginalComment = true;
            tocNcx.OptionOutputOriginalCase = true;
            tocNcx.LoadHtml($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><ncx xmlns=\"http://www.daisy.org/z3986/2005/ncx/\" version=\"2005-1\"><head><meta name=\"dtb:uid\" content=\"{BookGUID}\"/><meta name=\"dtb:totalPageCount\" content=\"0\"/><meta name=\"dtb:maxPageNumber\" content=\"0\"/></head><docTitle><text>{BookName}</text></docTitle><navMap></navMap></ncx>");

            
            var navMapNode = tocNcx.DocumentNode.SelectSingleNode("//navmap");
            for (int i = 1; i <= Directory.GetFiles(directory, "*.txt").Length; i++)
            {
                navMapNode.AppendChild(CreateNavPointNode(tocNcx, i));
            }

            ZipArchiveEntry tocNcxEntry = epub.CreateEntry("OEBPS/toc.ncx");

            using (MemoryStream buffer = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(buffer))
            using (StreamReader reader = new StreamReader(buffer))
            using (StreamWriter Outwriter = new StreamWriter(tocNcxEntry.Open()))
            {
                tocNcx.Save(writer);
                buffer.Flush();
                buffer.Position = 0;

                var XML = reader.ReadToEnd();
                XML = XML.Replace("<?xml version=\"1.0\" encoding=\"utf-8\"?><span>", "");
                XML = XML.Replace("</span>", "");

                Outwriter.Write(XML);
            }
        }

        static HtmlNode CreateManifestItemNode(HtmlDocument document, string fileName, string mediaType = "application/xhtml+xml", string ID = null)
        {
            var itemNode = document.CreateElement("item");
            itemNode.Attributes.Add("href", fileName);
            itemNode.Attributes.Add("media-type", mediaType);
            itemNode.Attributes.Add("id", ID ?? Path.GetFileNameWithoutExtension(fileName));
            return itemNode;
        }

        static HtmlNode CreateSpineItemNode(HtmlDocument document, string fileName)
        {
            var itemRefNode = document.CreateElement("itemref");
            itemRefNode.Attributes.Add("idref", Path.GetFileNameWithoutExtension(fileName));
            return itemRefNode;
        }

        static HtmlNode CreateNavPointNode(HtmlDocument document, int chapterNumber)
        {
            // Cria um novo elemento "navPoint"
            HtmlNode navPointNode = document.CreateElement("navPoint"); 

            // Adiciona os atributos "id" e "playOrder" ao elemento "navPoint"
            navPointNode.Attributes.Add("id", "chapter_" + chapterNumber.ToString());
            navPointNode.Attributes.Add("playOrder", chapterNumber.ToString());

            // Cria um novo elemento "navLabel" e adiciona ao elemento "navPoint"
            HtmlNode navLabelNode = document.CreateElement("navLabel");

            navPointNode.ChildNodes.Add(navLabelNode);

            // Cria um novo elemento "text", define seu conteúdo e adiciona ao elemento "navLabel"
            HtmlNode textNode = document.CreateElement("text");
            textNode.InnerHtml = "Chapter " + chapterNumber.ToString();

            navLabelNode.ChildNodes.Add(textNode);

            // Cria um novo elemento "content", adiciona o atributo "src" e adiciona ao elemento "navPoint"
            HtmlNode contentNode = document.CreateElement("content");

            contentNode.Attributes.Add("src", "chapter_" + chapterNumber.ToString() + ".html");
            navPointNode.ChildNodes.Add(contentNode);

            // Retorna o elemento "navPoint" criado
            return navPointNode;
        }

        static void CreateNamespace(HtmlDocument document)
        {
            // Obtém o elemento raiz do documento
            HtmlNode root = document.DocumentNode.SelectSingleNode("/");

            root.SetAttributeValue("xmlns:opf", "http://www.idpf.org/2007/opf");

        }

        static void SetOPFMetadata(HtmlDocument document)
        {
            var Metadata = document.DocumentNode.SelectSingleNode("/package/metadata");

            Metadata.SetAttributeValue("xmlns:dc", "http://purl.org/dc/elements/1.1/");
            Metadata.SetAttributeValue("xmlns:dcterms", "http://purl.org/dc/terms/");
            Metadata.SetAttributeValue("xmlns:opf", "http://www.idpf.org/2007/opf");
            Metadata.SetAttributeValue("xmlns:ncx", "http://www.daisy.org/z3986/2005/ncx/");
            Metadata.SetAttributeValue("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");

            var identifier = document.CreateElement("dc:identifier");
            identifier.Id = "uuid_id";
            identifier.SetAttributeValue("opf:scheme", "uuid");
            identifier.InnerHtml = BookGUID;

            var cover = document.CreateElement("meta");
            cover.SetAttributeValue("name", "cover");
            cover.SetAttributeValue("content", "img_cover");

            Metadata.AppendChild(identifier);
            Metadata.AppendChild(cover);
        }
    }
}