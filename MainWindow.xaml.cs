using ChangeLayoutStyle.Properties;
using Kompas6API5;
using KompasAPI7;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ChangeLayoutStyle
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private CancellationTokenSource tokenSource;
        private CancellationToken token;
        public MainWindow()
        {
            InitializeComponent();

            #region Загрузка настроек
            tb_folderDir.Text = Settings.Default.FolderDir;
            tb_layoutLibraryFileName.Text = Settings.Default.LayoutLibraryFileName;
            cb_dirs.IsChecked = Settings.Default.isDirs;
            #endregion
        }

        public List<string> Log { get => _log; set => _log = value; }

        private List<string> _log = new List<string>();

        private async Task ChangeLayoutAsync(string[] FilesDirs, string layoutLibraryFileName, string layoutStyleNumber, IProgress<int> progress)
        {
            progress.Report(10);
            Type kompasType = Type.GetTypeFromProgID("KOMPAS.Application.5", true);
            KompasObject kompas = Activator.CreateInstance(kompasType) as KompasObject; //Запуск компаса
            kompas.Visible = false;
            int progressCount = 20;
            progress.Report(progressCount);
            IApplication application = (IApplication)kompas.ksGetApplication7();
            if (token.IsCancellationRequested)
            {
                application.Quit();
                progress.Report(0);
                return;
            }
            IDocuments documets = application.Documents;
            foreach (string item in FilesDirs)
            {
                IKompasDocument kompasDocument = documets.Open(item, false, false);
                if (kompasDocument is null)
                {
                    Log.Add($"{item} - Не получилось открыть документ");
                    break;
                }

                ILayoutSheets layoutSheets = kompasDocument.LayoutSheets;
                if (layoutSheets.Count == 0)
                {
                    Log.Add($"{item} - Листов нет");
                    break;
                }
                ILayoutSheet layoutSheet = null;
                foreach (ILayoutSheet item1 in layoutSheets)
                {
                    layoutSheet = item1;
                    break;
                }
                if (layoutSheet == null)
                {
                    Log.Add(item);
                }
                layoutSheet.LayoutLibraryFileName = layoutLibraryFileName;
                layoutSheet.LayoutStyleNumber = Convert.ToDouble(layoutStyleNumber);
                layoutSheet.Update();
                kompasDocument.Save();
                if (kompasDocument.Changed)
                {
                    kompasDocument.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                    Log.Add($"{item} - не сохранен");
                }
                kompasDocument.Close(Kompas6Constants.DocumentCloseOptions.kdSaveChanges);
                progressCount += 80 / FilesDirs.Length;
                progress.Report(progressCount);
                if (token.IsCancellationRequested)
                {
                    application.Quit();
                    progress.Report(0);
                    return;
                }
            }
            application.Quit();
            progress.Report(100);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                InitialDirectory = tb_folderDir.Text,
                IsFolderPicker = true
            };
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                tb_folderDir.Text = dialog.FileName;
            }
        }

        private void b_layoutLibraryFileName_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                DefaultFileName = "*.lyt"
            };
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                tb_layoutLibraryFileName.Text = dialog.FileName;
            }
        }

        private async void b_change_Click(object sender, RoutedEventArgs e)
        {
            b_Cancel.IsEnabled = true;
            b_change.IsEnabled = false;
            Log.Clear();
            if (!Directory.Exists(tb_folderDir.Text))
            {
                tb_finish.Text = "Путь к папке  с файлами не корректен.";
                return;
            }
            if (!File.Exists(tb_layoutLibraryFileName.Text))
            {
                tb_finish.Text = "Путь к файлу оформления не корректен.";
                return;
            }
            if (tb_layoutLibraryFileName.Text == "" || tb_LayoutStyleNumber.Text == "")
            {
                tb_finish.Text = "Нет данных для изменения.";
                return;
            }

            string[] FilesDirs = new string[0];
            if (cb_dirs.IsChecked == true)
            {
                FilesDirs = Directory.GetFiles(tb_folderDir.Text, "*.cdw", SearchOption.AllDirectories);
            }
            else
            {
                FilesDirs = Directory.GetFiles(tb_folderDir.Text, "*.cdw");
            }
            string layoutLibraryFileName = tb_layoutLibraryFileName.Text;
            string layoutStyleNumber = tb_LayoutStyleNumber.Text;
            tb_finish.Text = "Началось изменение";
            var progress = new Progress<int>( value =>
                {
                    progressbar.Value = value;
                });
            tokenSource = new CancellationTokenSource();
            token = tokenSource.Token;
            await Task.Run(() => ChangeLayoutAsync(FilesDirs, layoutLibraryFileName, layoutStyleNumber, progress), token);
            b_change.IsEnabled = true;
            b_Cancel.IsEnabled = false;
            if (token.IsCancellationRequested)
            {
                tb_finish.Text = "Отменено";
                return;
            }

            if (Log.Count == 0)
            {
                tb_finish.Text = "Готово";
                return;
            }
            using (StreamWriter sw = new StreamWriter("Log.txt", false))
            {
                foreach (var item in Log)
                {
                    sw.WriteLine(item);
                }
                sw.Close();
            }
            tb_finish.Text = "Готово. Часть файлов не была изменена, просмотрите журнал.";


        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            #region Сохранение настроек
            Settings.Default.FolderDir = tb_folderDir.Text;
            Settings.Default.LayoutLibraryFileName = tb_layoutLibraryFileName.Text;
            Settings.Default.isDirs = (bool)cb_dirs.IsChecked;
            Settings.Default.Save();
            #endregion
        }

        private void b_log_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists("Log.txt"))
            {
                Process.Start("Log.txt");
            }
            else
            {
                tb_finish.Text = "Файл журнала не найден.";
            }
            
        }

        private void b_Cancel_Click(object sender, RoutedEventArgs e)
        {
            tokenSource.Cancel();
        }
    }
}
