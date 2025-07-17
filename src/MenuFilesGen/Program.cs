using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace MenuFilesGen
{
    internal static class Program
    {
        public static XElement CreateButton(string[] commandData)
        {
            var ribbonCommandButton = new XElement("RibbonCommandButton");
            ribbonCommandButton.Add(new XAttribute("Text", commandData[0]));
            ribbonCommandButton.Add(new XAttribute("ButtonStyle", commandData[4]));
            ribbonCommandButton.Add(new XAttribute("MenuMacroID", commandData[1]));
            return ribbonCommandButton;
        }

        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main(string[] args)
        {
            OpenFileDialog tableFileDialog = new OpenFileDialog
            {
                Filter = "TSV files (*.tsv)|*.tsv",
            };

            if (tableFileDialog.ShowDialog() != DialogResult.OK)
                return;

            var directoryPath = Path.GetDirectoryName(tableFileDialog.FileName);
            var addinName = Path.GetFileNameWithoutExtension(tableFileDialog.FileName);

            // Описания команд,сгруппированных по имени панели
            List<IGrouping<string, string[]>> commands;
            using (StreamReader reader = new StreamReader(tableFileDialog.FileName))
            {
                commands = reader
                    .ReadToEnd()
                    .Split('\n')
                    .Skip(1) // Заголовок таблицы
                    .Select(c => c.Split('\t')) // Разделитель - табуляция
                    .Where(c => !(c.Count() > 6 && c[6] == "TRUE")) // Пропуск скрытых команд
                    .GroupBy(c => c[3])
                    .ToList();
            }

            var cfgFilePath = $"{directoryPath}\\{addinName}.cfg";
            var cuiFilePath = $"{directoryPath}\\RibbonRoot.cui";
            var cuixFilePath = $"{directoryPath}\\{addinName}.cuix";

            using (StreamWriter writer = new StreamWriter(cfgFilePath, false, new UTF8Encoding(true)))
            {
                // Регистрация команд
                writer.WriteLine(
                    $"[\\ribbon\\{addinName}]"
                    + $"\r\nCUIX=s%CFG_PATH%\\{addinName}.cuix"
                    + "\r\n"
                    + "\r\n[\\configman]"
                    + "\r\n[\\configman\\commands]");

                foreach (IGrouping<string, string[]> commandGroup in commands)
                {
                    foreach (string[] commandData in commandGroup)
                    {
                        writer.WriteLine(
                            $@"[\configman\commands\{commandData[1]}]"
                            + $"\r\nweight=i10"
                            + $"\r\ncmdtype=i1"
                            + $"\r\nintername=s{commandData[1]}"
                            + $"\r\nDispName=s{commandData[0]}"
                            + $"\r\nStatusText=s{commandData[2]}"
                            + $"\r\nBitmapDll=sicons\\{commandData[1]}.ico");
                    }
                }

                // Классическое меню
                writer.WriteLine("\r\n[\\menu]" + $"\r\n[\\menu\\{addinName}_Menu]" + $"\r\nName=s{addinName}");

                foreach (IGrouping<string, string[]> commandGroup in commands)
                {
                    writer.WriteLine($@"[\menu\{addinName}_Menu\{commandGroup.Key}]" + $"\r\nname=s{commandGroup.Key}");

                    foreach (string[] commandData in commandGroup)
                    {
                        writer.WriteLine(
                            $@"[\menu\{addinName}_Menu\{commandGroup.Key}\s{commandData[1]}]"
                            + $"\r\nname=s{commandData[0]}"
                            + $"\r\nIntername=s{commandData[1]}");
                    }
                }

                // Панели инструментов
                writer.WriteLine("\r\n[\\toolbars]");

                foreach (IGrouping<string, string[]> commandGroup in commands)
                {
                    var panelName = $"{addinName}_{commandGroup.Key}";
                    writer.WriteLine(
                        $"[\\toolbars\\{panelName}]" + $"\r\nname=s{panelName}" + $"\r\nIntername=s{panelName}");

                    foreach (string[] commandData in commandGroup)
                    {
                        writer.WriteLine(
                            $"[\\toolbars\\{panelName}\\{commandData[1]}]" + $"\r\nIntername=s{commandData[1]}");
                    }
                }
            }

            // Ленточное меню
            //Создание XML документа 
            var xDoc = new XDocument();

            //Корневой элемент
            var ribbonRoot = new XElement("RibbonRoot");
            xDoc.Add(ribbonRoot);

            var ribbonPanelSourceCollection = new XElement("RibbonPanelSourceCollection");
            ribbonRoot.Add(ribbonPanelSourceCollection);

            var ribbonTabSourceCollection = new XElement("RibbonTabSourceCollection");
            ribbonRoot.Add(ribbonTabSourceCollection);

            var ribbonTabSource = new XElement("RibbonTabSource");
            ribbonTabSource.Add(new XAttribute("Text", addinName));
            ribbonTabSource.Add(new XAttribute("UID", $"{addinName.Replace(" ", "")}_Tab"));
            ribbonTabSourceCollection.Add(ribbonTabSource);

            foreach (IGrouping<string, string[]> commandGroup in commands)
            {
                var ribbonPanelSource = new XElement("RibbonPanelSource");
                ribbonPanelSource.Add(new XAttribute("UID", commandGroup.Key));
                ribbonPanelSource.Add(new XAttribute("Text", commandGroup.Key));
                ribbonPanelSourceCollection.Add(ribbonPanelSource);

                // Временный контейнер для сбора кнопок
                var panelButtons = new XElement("Temp");

                // Группируем команды по тому, объединяются ли они в RibbonSplitButton
                var unitedCommands = commandGroup.GroupBy(c => c[5]).ToList();

                foreach (var unitedCommandGroup in unitedCommands)
                {
                    XElement container = panelButtons;
                    if (!string.IsNullOrWhiteSpace(unitedCommandGroup.Key))
                    {
                        var ribbonSplitButton = new XElement("RibbonSplitButton");
                        ribbonSplitButton.Add(new XAttribute("Text", unitedCommandGroup.Key));
                        ribbonSplitButton.Add(new XAttribute("Behavior", "SplitFollowStaticText"));
                        ribbonSplitButton.Add(new XAttribute("ButtonStyle", unitedCommandGroup.First()[4]));

                        panelButtons.Add(ribbonSplitButton);
                        container = ribbonSplitButton;
                    }

                    foreach (string[] commandData in unitedCommandGroup)
                        container.Add(CreateButton(commandData));
                }

                var sortedButtons = panelButtons
                    .Elements()
                    .GroupBy(
                        button => button.Attributes().First(attr => attr.Name == "ButtonStyle").Value.Contains("Small"))
                    .ToDictionary(g => g.Key);

                if (sortedButtons.ContainsKey(false))
                {
                    foreach (var button in sortedButtons[false])
                        ribbonPanelSource.Add(button);
                }

                if (sortedButtons.ContainsKey(true))
                {
                    XElement ribbonRowPanel = null;
                    var ribbonRowPanelButtonsCount = 3;

                    foreach (var button in sortedButtons[true])
                    {
                        if (ribbonRowPanelButtonsCount == 3)
                        {
                            ribbonRowPanel = new XElement("RibbonRowPanel");
                            ribbonPanelSource.Add(ribbonRowPanel);
                            ribbonRowPanelButtonsCount = 0;
                        }

                        var ribbonRow = new XElement("RibbonRow");
                        ribbonRow.Add(button);
                        ribbonRowPanel.Add(ribbonRow);

                        ribbonRowPanelButtonsCount++;
                    }
                }

                // Дублирование
                XElement ribbonPanelBreak = new XElement("RibbonPanelBreak");
                ribbonPanelSource.Add(ribbonPanelBreak);
                var ribbonRowDuplicatePanel = new XElement("RibbonRowPanel");
                ribbonPanelSource.Add(ribbonRowDuplicatePanel);

                var items = panelButtons.Elements().ToArray();
                int nameSymbolsCountMax = 0;

                for (int itemIndex = 0; itemIndex < items.Count(); itemIndex += 2)
                {
                    var nameSymbolsCount =
                        items[itemIndex].Attributes().First(attr => attr.Name == "Text").Value.Count();

                    if (nameSymbolsCount > nameSymbolsCountMax)
                        nameSymbolsCountMax = nameSymbolsCount;
                }

                for (int itemIndex = 0; itemIndex < items.Count(); itemIndex++)
                {
                    var item = items[itemIndex];
                    XElement[] itemButtons;

                    if (item.Name == "RibbonSplitButton")
                        itemButtons = item.Elements().ToArray();
                    else
                    {
                        itemButtons = new[]
                        {
                            item,
                        };
                    }

                    for (int buttonIndex = 0; buttonIndex < itemButtons.Count(); buttonIndex++)
                    {
                        var button = itemButtons[buttonIndex];
                        button.Attributes().First(attr => attr.Name == "ButtonStyle").Value = "LargeWithHorizontalText";
                        ribbonRowDuplicatePanel.Add(button);

                        if (itemIndex < items.Count() - 1 || buttonIndex < itemButtons.Count() - 1)
                        {
                            XElement separator = new XElement("RibbonSeparator");
                            ribbonRowDuplicatePanel.Add(separator);
                        }
                    }
                }

                var ribbonPanelSourceReference = new XElement("RibbonPanelSourceReference");
                ribbonPanelSourceReference.Add(new XAttribute("PanelId", commandGroup.Key));
                ribbonTabSource.Add(ribbonPanelSourceReference);
            }

            xDoc.Save(cuiFilePath);

            // Удаление .cuix файла, если он существует
            if (File.Exists(cuixFilePath))
                File.Delete(cuixFilePath);

            // Создание архива (.cuix), добавление в него сформированного .xml файла
            using (var zip = ZipFile.Open(cuixFilePath, ZipArchiveMode.Create))
                zip.CreateEntryFromFile(cuiFilePath, "RibbonRoot.cui");

            MessageBox.Show($"Файлы {addinName}.cfg и {addinName}.cuix сохранены в папке {directoryPath}");
        }
    }
}