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
            cb_redactionAuto.IsChecked = Settings.Default.isRedactionAuto;
            #endregion
        }

        public List<string> Log { get => _log; set => _log = value; }

        private List<string> _log = new List<string>();

        /// <summary>
        /// Асинхронный метод изменения оформления
        /// </summary>
        /// <param name="FilesDirs"></param>
        /// <param name="layoutLibraryFileName"></param>
        /// <param name="layoutStyleNumber"></param>
        /// <param name="progress"></param>
        /// <returns></returns>
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
        /// <summary>
        /// Выбор папки с чертежами
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Выбор файла оформления
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Изменение оформления
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void b_change_Click(object sender, RoutedEventArgs e)
        {
            b_Cancel.IsEnabled = true;
            b_change.IsEnabled = false;
            Log.Clear();
            if (!Directory.Exists(tb_folderDir.Text))
            {
                tb_finish.Text = "Путь к папке  с файлами не корректен.";
                b_change.IsEnabled = true;
                b_Cancel.IsEnabled = false;
                return;
            }
            if (!File.Exists(tb_layoutLibraryFileName.Text))
            {
                tb_finish.Text = "Путь к файлу оформления не корректен.";
                b_change.IsEnabled = true;
                b_Cancel.IsEnabled = false;
                return;
            }
            if (tb_layoutLibraryFileName.Text == "" || tb_LayoutStyleNumber.Text == "")
            {
                tb_finish.Text = "Нет данных для изменения.";
                b_change.IsEnabled = true;
                b_Cancel.IsEnabled = false;
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
            #region Для вызова асинхронного метода изменения оформления
            string layoutLibraryFileName = tb_layoutLibraryFileName.Text;
            string layoutStyleNumber = tb_LayoutStyleNumber.Text;
            tb_finish.Text = "Началось изменение";
            var progress = new Progress<int>(value =>
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
            #endregion

            using (StreamWriter sw = new StreamWriter("Log.txt", false))
            {
                foreach (var item in Log)
                {
                    sw.WriteLine(item);
                }
                sw.Close();
            }
            if (Log.Count == 0)
            {
                tb_finish.Text = "Готово";
            }
            else
            {
                tb_finish.Text = "Готово. Часть файлов не была изменена, просмотрите журнал.";
            }


        }
        /// <summary>
        /// Закрыть окно и сохранить настройки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            #region Сохранение настроек
            Settings.Default.FolderDir = tb_folderDir.Text;
            Settings.Default.LayoutLibraryFileName = tb_layoutLibraryFileName.Text;
            Settings.Default.isDirs = (bool)cb_dirs.IsChecked;
            Settings.Default.isRedactionAuto = (bool)cb_redactionAuto.IsChecked;
            Settings.Default.Save();
            #endregion
        }
        /// <summary>
        /// Открыть файл журнала
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Отмена процесса изменения оформления
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void b_Cancel_Click(object sender, RoutedEventArgs e)
        {
            tokenSource.Cancel();
        }
        /// <summary>
        /// Добавление текущей даты
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void b_DataRedaction_Click(object sender, RoutedEventArgs e)
        {
            tb_DataRedaction.Text = DateTime.Now.ToString("dd.MM");
        }
        /// <summary>
        /// Создание редакции
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void b_Redaction_Click(object sender, RoutedEventArgs e)
        {
            b_Redaction_Cancel.IsEnabled = true;
            b_Redaction.IsEnabled = false;
            Log.Clear();
            if (!Directory.Exists(tb_folderDir.Text))
            {
                tb_finish.Text = "Путь к папке  с файлами не корректен.";
                b_Redaction.IsEnabled = true;
                b_Redaction_Cancel.IsEnabled = false;
                return;
            }
            if (tb_NumberRedaction.Text == "" || tb_DataRedaction.Text == "")
            {
                tb_finish.Text = "Нет данных для изменения.";
                b_Redaction.IsEnabled = true;
                b_Redaction_Cancel.IsEnabled = false;
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
            #region Для вызова асинхронного метода изменения оформления
            string numberRedaction = tb_NumberRedaction.Text;
            string dataRedaction = tb_DataRedaction.Text;
            bool? isAuto = cb_redactionAuto.IsChecked;
            tb_finish.Text = "Началось изменение";
            var progress = new Progress<int>(value =>
            {
                progressbar.Value = value;
            });
            tokenSource = new CancellationTokenSource();
            token = tokenSource.Token;
            await Task.Run(() => SetRedaction(FilesDirs, numberRedaction, dataRedaction, progress, isAuto), token);
            b_Redaction.IsEnabled = true;
            b_Redaction_Cancel.IsEnabled = false;
            if (token.IsCancellationRequested)
            {
                tb_finish.Text = "Отменено";
                return;
            }
            #endregion

            using (StreamWriter sw = new StreamWriter("Log.txt", false))
            {
                foreach (var item in Log)
                {
                    sw.WriteLine(item);
                }
                sw.Close();
            }
            if (Log.Count == 0)
            {
                tb_finish.Text = "Готово";
            }
            else
            {
                tb_finish.Text = "Готово. Часть файлов не была изменена, просмотрите журнал.";
            }
        }
        /// <summary>
        /// Создание редакции
        /// </summary>
        /// <param name="FilesDirs"></param>
        /// <param name="numberRedaction"></param>
        /// <param name="dataRedaction"></param>
        /// <param name="progress"></param>
        /// <param name="iSAuto"></param>
        /// <returns></returns>
        private async Task SetRedaction(string[] FilesDirs, string numberRedaction, string dataRedaction, IProgress<int> progress, bool? iSAuto)
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
                IStamp stamp = layoutSheet.Stamp;
                List<string[]> redaction = new List<string[]>();
                for (int i = 0; i < 4; i++)
                {
                    if (stamp.Text[140 + i].Str != "")
                    {
                        redaction.Add( new string[]
                        {
                            stamp.Text[140 + i].Str,
                            stamp.Text[150 + i].Str,
                            stamp.Text[180 + i].Str
                        });
                    }
                }
                if (iSAuto == true)
                {
                    int numberRedactionInt;
                    int.TryParse((redaction[redaction.Count -1 ][0]), out numberRedactionInt);
                    numberRedactionInt += 1;
                    numberRedaction = numberRedactionInt.ToString();
                }
                redaction.Add(new string[] {numberRedaction, "ред.", dataRedaction});
                int increment = 0;
                if (redaction.Count == 5) { increment = 1; }
                for (int i = 0; i < redaction.Count - increment; i++) 
                {
                    stamp.Text[140 + i].Str = redaction[i + increment][0];
                    stamp.Text[150 + i].Str = redaction[i + increment][1];
                    stamp.Text[180 + i].Str = redaction[i + increment][2];
                }
                stamp.Update();
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


        private async void b_ClearCell_Click(object sender, RoutedEventArgs e)
        {
            b_Redaction_Cancel.IsEnabled = true;
            b_ClearCell.IsEnabled = false;
            Log.Clear();
            if (!Directory.Exists(tb_folderDir.Text))
            {
                tb_finish.Text = "Путь к папке  с файлами не корректен.";
                b_ClearCell.IsEnabled = true;
                b_Redaction_Cancel.IsEnabled = false;
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
            #region Для вызова асинхронного метода изменения оформления
            int[] numberCels = new int[]
            {
                140, 141, 142, 143,
                150, 151, 152, 153,
                160, 161, 162, 163,
                170, 171, 172, 173,
                180, 181, 182, 183
            };
            tb_finish.Text = "Началось изменение";
            var progress = new Progress<int>(value =>
            {
                progressbar.Value = value;
            });
            tokenSource = new CancellationTokenSource();
            token = tokenSource.Token;
            await Task.Run(() => ClearStamp(FilesDirs, numberCels, progress), token);
            b_Redaction.IsEnabled = true;
            b_Redaction_Cancel.IsEnabled = false;
            if (token.IsCancellationRequested)
            {
                tb_finish.Text = "Отменено";
                return;
            }
            #endregion

            using (StreamWriter sw = new StreamWriter("Log.txt", false))
            {
                foreach (var item in Log)
                {
                    sw.WriteLine(item);
                }
                sw.Close();
            }
            if (Log.Count == 0)
            {
                tb_finish.Text = "Готово";
            }
            else
            {
                tb_finish.Text = "Готово. Часть файлов не была изменена, просмотрите журнал.";
            }
        }
        /// <summary>
        /// Очистка ячек штампа
        /// </summary>
        /// <param name="numberCells"></param>
        private async void ClearStamp(string[] FilesDirs, int[] numberCells, IProgress<int> progress)
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
                IStamp stamp = layoutSheet.Stamp;
                foreach (var nuberCell in numberCells)
                {
                    stamp.Text[nuberCell].Clear();
                }
                stamp.Update();
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
        private void cb_redactionAuto_Click(object sender, RoutedEventArgs e)
        {
            if (cb_redactionAuto.IsChecked == true)
            {
                tb_NumberRedaction.IsEnabled = false;
            }
            if (cb_redactionAuto.IsChecked == false)
            {
                tb_NumberRedaction.IsEnabled = true;
            }
        }

        private void cb_redactionAuto_Checked(object sender, RoutedEventArgs e)
        {
            if (cb_redactionAuto.IsChecked == true)
            {
                tb_NumberRedaction.IsEnabled = false;
            }
            if (cb_redactionAuto.IsChecked == false)
            {
                tb_NumberRedaction.IsEnabled = true;
            }
        }

        private void b_Redaction_Cancel_Click(object sender, RoutedEventArgs e)
        {
            tokenSource.Cancel();
        }

    }
}
