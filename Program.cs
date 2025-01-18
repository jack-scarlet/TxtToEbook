using HtmlAgilityPack;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
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
        static string[] chapters;

        static void Main(string[] args)
        {
            if (args.Length != 1)
                return;

            directory = args.First();
            outputFileName = args.Last().TrimEnd(' ', '/', '\\') + ".epub";

            Console.WriteLine("Digite o nome do eBook:");
            BookName = Console.ReadLine() ?? throw new Exception();
            Console.WriteLine("Digite o nome do Autor:");
            BookAuthor = Console.ReadLine() ?? throw new Exception();

            static void ProcessBook(string Book)
            {
                var Lines = File.ReadAllLines(Book);

                if (Lines.Length == 0)
                    throw new Exception($"O arquivo {Path.GetFileName(Book)} está vazio.");

                var Trimmer = (string x) => x.Trim().Trim('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');
                var DummyLine = (string x) => x.Length > 0 && x.Trim(x.Trim().First(), ' ', '\t') == string.Empty;

                var ChapterPrefix = Lines.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
                if (ChapterPrefix == null || !char.IsLetter(ChapterPrefix.Trim().First()))
                    throw new Exception($"O arquivo {Path.GetFileName(Book)} está mal formatado e não contém capítulos válidos.");

                ChapterPrefix = Trimmer(ChapterPrefix);

                List<string> ChapterContent = new List<string>();
                int chp = 0; // Inicializa a variável chp
                for (int i = 0; i < Lines.Length; i++)
                {
                    var Line = Lines[i];
                    var Prefix = Trimmer(Line);

                    if (Prefix == ChapterPrefix)
                    {
                        if (ChapterContent.Count > 0 || chp > 0)
                        {
                            string SaveAs = Path.Combine(directory, $"{chp}.txt");
                            File.WriteAllLines(SaveAs, ChapterContent.ToArray());
                            ChapterContent.Clear();
                        }

                        chp = int.Parse(Line.Replace(Prefix, "").Trim());
                        continue;
                    }

                    if (Line.StartsWith("<image>"))
                    {
                        string imageFileName = Line.Substring(7).Trim();
                        ChapterContent.Add($"<p><img src=\"img/{imageFileName}\" alt=\"Chapter {chp} Illustration\"/></p>");
                        continue;
                    }

                    // Adiciona suporte para cabeçalhos
                    if (Line.StartsWith("<h1>") || Line.StartsWith("<h2>") || Line.StartsWith("<h3>") || Line.StartsWith("<h4>") || Line.StartsWith("<h5>") || Line.StartsWith("<h6>"))
                    {
                        int headerLevel = int.Parse(Line.Substring(2, 1));
                        string headerContent = Line.Substring(4).Trim();
                        ChapterContent.Add($"<h{headerLevel}>{HttpUtility.HtmlEncode(headerContent)}</h{headerLevel}>");
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(Line))
                    {
                        ChapterContent.Add("<br />");
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(Line))
                    {
                        if (Line.StartsWith("**"))
                            Line = " " + Line;

                        if (Line.EndsWith("**"))
                            Line += " ";

                        ChapterContent.Add($"<p>{HttpUtility.HtmlEncode(Line).Replace(" **", "<i>").Replace("** ", "</i>")}</p>");
                    }
                }

                if (ChapterContent.Count > 0 || chp > 0)
                {
                    string SaveAs = Path.Combine(directory, $"{chp}.txt");
                    File.WriteAllLines(SaveAs, ChapterContent.ToArray());
                }
            }


            if (Directory.GetFiles(directory, "*.btxt").Length > 0)
            {
                var Files = Directory.GetFiles(directory, "*.btxt");
                foreach (var Book in Files)
                {
                    ProcessBook(Book);
                }
            }

            // Criar diretório temporário
            Directory.CreateDirectory(tempDirectory);

            // Separar conteúdo em capítulos
            chapters = Directory.GetFiles(directory, "*.txt").OrderBy(x => int.Parse(Path.GetFileNameWithoutExtension(x))).ToArray();

            // Processar cada arquivo TXT
            for (int i = 0; i < chapters.Length; i++)
            {
                var file = chapters[i];
                string[] lines = File.ReadAllLines(file);

                bool onlyImage = lines.Length == 1 && lines[0].Contains("<img");

                // Verifica se a primeira linha possui uma tag de cabeçalho
                string chapterTitle = null;
                if (lines.Length > 0 &&
                    (lines[0].StartsWith("<h1>") || lines[0].StartsWith("<h2>") ||
                     lines[0].StartsWith("<h3>") || lines[0].StartsWith("<h4>") ||
                     lines[0].StartsWith("<h5>") || lines[0].StartsWith("<h6>")))
                {
                    // Extraia o texto dentro da tag de cabeçalho
                    chapterTitle = Regex.Replace(lines[0], @"<[^>]+>", "").Trim();
                }

                // Criar conteúdo HTML
                StringBuilder htmlContent = new StringBuilder();
                htmlContent.Append("<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>");
                if (chapterTitle != null)
                {
                    htmlContent.Append(chapterTitle);
                }
                else
                {
                    htmlContent.Append("Chapter " + (i + 1));
                }
                htmlContent.Append("</title></head><body>");

                if (!onlyImage)
                {
                    if (chapterTitle != null)
                    {
                        htmlContent.Append(lines[0]); // Adicionar a linha completa do cabeçalho
                        // Remover a linha do cabeçalho do corpo do HTML
                        lines = lines.Skip(1).ToArray();
                    }
                    else
                    {
                        htmlContent.Append("<h1>Chapter " + (i + 1) + "</h1>");
                    }
                }

                htmlContent.Append(string.Join("\n", lines));
                htmlContent.Append("</body></html>");

                // Criar arquivo HTML para o capítulo
                string chapterHtmlPath = Path.Combine(tempDirectory, $"chapter_{i + 1}.html");
                File.WriteAllText(chapterHtmlPath, htmlContent.ToString());
            }

            // Remover arquivo EPUB existente
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

                // Adicionar conteúdo HTML
                foreach (string file in Directory.GetFiles(tempDirectory, "*.html").OrderBy(x => int.Parse(Path.GetFileNameWithoutExtension(x).Replace("chapter_", ""))))
                {
                    epub.CreateEntryFromFile(file, $"OEBPS/{Path.GetFileName(file)}");
                }

                // Adicionar imagens
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

            // Criar página de capa
            string coverPage = "<?xml version='1.0' encoding='utf-8'?><html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:lang=\"en\"><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"/><meta name=\"calibre:cover\" content=\"true\"/><title>Cover</title><style type=\"text/css\" title=\"override_css\">@page {padding: 0pt; margin:0pt} body { text-align: center; padding:0pt; margin: 0pt; }</style></head><body><div><img width=\"1200\" height=\"1600\" src=\"img/cover.jpg\" alt=\"Book Cover\"/></div></body></html>";
            ZipArchiveEntry coverEntry = epub.CreateEntry("OEBPS/cover.xhtml");
            using (StreamWriter writer = new StreamWriter(coverEntry.Open()))
            {
                writer.Write(coverPage);
            }

            // Criar documento content.opf
            var contentOpf = new HtmlDocument();
            contentOpf.OptionOutputAsXml = true;
            contentOpf.OptionPreserveXmlNamespaces = true;
            contentOpf.LoadHtml($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><package xmlns=\"http://www.idpf.org/2007/opf\" version=\"2.0\" unique-identifier=\"uuid_id\"><metadata><dc:title>{BookName}</dc:title><dc:creator>{BookAuthor}</dc:creator><dc:language>en</dc:language><dc:publisher>{BookAuthor}</dc:publisher><dc:rights>Creative Commons</dc:rights></metadata><manifest></manifest><spine toc=\"ncx\"></spine></package>");

            CreateNamespace(contentOpf);
            SetOPFMetadata(contentOpf);

            // Encontrar nós manifest e spine
            var manifestNode = contentOpf.DocumentNode.SelectSingleNode("//manifest");
            if (manifestNode == null)
                throw new NullReferenceException("O nó 'manifest' não foi encontrado no arquivo content.opf.");

            var spineNode = contentOpf.DocumentNode.SelectSingleNode("//spine");
            if (spineNode == null)
                throw new NullReferenceException("O nó 'spine' não foi encontrado no arquivo content.opf.");

            // Adicionar arquivos HTML ao manifesto e espinha dorsal
            var htmlFiles = Directory.GetFiles(tempDirectory, "*.html").OrderBy(x => int.Parse(Path.GetFileNameWithoutExtension(x).Replace("chapter_", "")));
            foreach (string file in htmlFiles)
            {
                string fileName = Path.GetFileName(file);
                manifestNode.AppendChild(CreateManifestItemNode(contentOpf, fileName));
                spineNode.AppendChild(CreateSpineItemNode(contentOpf, fileName));
            }

            // Adicionar imagens ao manifesto
            var imageFiles = Directory.GetFiles(directory, "*.jpg", SearchOption.AllDirectories);
            foreach (var image in imageFiles)
            {
                manifestNode.AppendChild(CreateManifestItemNode(contentOpf, $"img/{Path.GetFileName(image)}", "image/jpeg", $"img_{Path.GetFileNameWithoutExtension(image)}"));
            }

            // Adicionar página de capa ao manifesto e espinha dorsal
            var coverImagePath = Path.Combine(directory, "cover.jpg");
            if (File.Exists(coverImagePath))
            {
                manifestNode.AppendChild(CreateManifestItemNode(contentOpf, "cover.xhtml"));
                spineNode.InsertBefore(CreateSpineItemNode(contentOpf, "cover.xhtml"), spineNode.FirstChild);
            }

            // Adicionar arquivo toc.ncx ao manifesto
            manifestNode.AppendChild(CreateManifestItemNode(contentOpf, "toc.ncx", "application/x-dtbncx+xml", "ncx"));

            // Criar entrada content.opf no EPUB
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

            static HtmlNode CreateNavPointNodeFromHeader(HtmlDocument document, HtmlNode headerNode, int playOrder, string fileName)
            {
                // Cria um novo elemento "navPoint"
                HtmlNode navPointNode = document.CreateElement("navPoint");

                // Adiciona os atributos "id" e "playOrder" ao elemento "navPoint"
                navPointNode.Attributes.Add("id", headerNode.Name + "_" + playOrder.ToString());
                navPointNode.Attributes.Add("playOrder", playOrder.ToString());

                // Cria um novo elemento "navLabel" e adiciona ao elemento "navPoint"
                HtmlNode navLabelNode = document.CreateElement("navLabel");
                navPointNode.ChildNodes.Add(navLabelNode);

                // Cria um novo elemento "text", define seu conteúdo e adiciona ao elemento "navLabel"
                HtmlNode textNode = document.CreateElement("text");
                textNode.InnerHtml = headerNode.InnerText;
                navLabelNode.ChildNodes.Add(textNode);

                // Cria um novo elemento "content", adiciona o atributo "src" e adiciona ao elemento "navPoint"
                HtmlNode contentNode = document.CreateElement("content");
                contentNode.Attributes.Add("src", "chapter_" + Path.GetFileNameWithoutExtension(fileName) + ".html");
                navPointNode.ChildNodes.Add(contentNode);

                // Retorna o elemento "navPoint" criado
                return navPointNode;
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
            var files = Directory.GetFiles(directory, "*.txt");
            int playOrder = 1;
            bool headersFound = false;

            foreach (var file in files)
            {
                var doc = new HtmlDocument();
                doc.Load(file);

                // Encontra todas as tags <h1> e <h2> (adapte conforme necessário)
                var headerNodes = doc.DocumentNode.SelectNodes("//h1 | //h2");

                if (headerNodes != null)
                {
                    headersFound = true; // Indica que encontramos headers
                    foreach (var headerNode in headerNodes)
                    {
                        navMapNode.AppendChild(CreateNavPointNodeFromHeader(tocNcx, headerNode, playOrder, file));
                        playOrder++;
                    }
                }
            }

            // Se não encontrarmos tags de cabeçalho, vamos criar o índice com base nos capítulos
            if (!headersFound)
            {
                for (int i = 1; i <= files.Length; i++)
                {
                    navMapNode.AppendChild(CreateNavPointNode(tocNcx, i));
                }
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
