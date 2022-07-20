using ChangeLayoutStyle.Properties;
using Kompas6API5;
using KompasAPI7;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace ChangeLayoutStyle
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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

        private bool ChangeLayout(IKompasDocument kompasDocument)
        {
            if (kompasDocument is null)
            {
                tb_finish.Text = "Не получилось открыть документ";
                return false;
            }
            
            ILayoutSheets layoutSheets = kompasDocument.LayoutSheets;
            if (layoutSheets.Count == 0)
            {
                tb_finish.Text = "Листов нет";
            }
            if (layoutSheets.Count > 1)
            {
                tb_finish.Text = "Листов больше одного";
            }
            ILayoutSheet layoutSheet = layoutSheets.ItemByNumber[1];
            if (tb_layoutLibraryFileName.Text != "")
            {
                layoutSheet.LayoutLibraryFileName = tb_layoutLibraryFileName.Text;
            }
            if (tb_LayoutStyleNumber.Text != "")
            {
                layoutSheet.LayoutStyleNumber = Convert.ToDouble(tb_LayoutStyleNumber.Text);
            }
            layoutSheet.Update();
            kompasDocument.Save();
            if (kompasDocument.Changed)
            {
                kompasDocument.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                return false;
            }
            kompasDocument.Close(Kompas6Constants.DocumentCloseOptions.kdSaveChanges);
            return true;
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

        private void b_change_Click(object sender, RoutedEventArgs e)
        {
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
            if (tb_layoutLibraryFileName.Text == "" && tb_LayoutStyleNumber.Text == "")
            {
                tb_finish.Text = "Нет данных для изменения.";
                return;
            }
            Type kompasType = Type.GetTypeFromProgID("KOMPAS.Application.5", true);
            KompasObject kompas = Activator.CreateInstance(kompasType) as KompasObject;
            kompas.Visible = false;

            IApplication application = (IApplication)kompas.ksGetApplication7();
            IDocuments documets = application.Documents;
            string[] FilesDirs = new string[0];
            if (cb_dirs.IsChecked == true)
            {
                FilesDirs = Directory.GetFiles(tb_folderDir.Text, "*.cdw", SearchOption.AllDirectories);
            }
            else
            {
                FilesDirs = Directory.GetFiles(tb_folderDir.Text, "*.cdw");
            }
            foreach (string item in FilesDirs)
            {
                IKompasDocument kompasDocument = documets.Open(item, false, false);
                if (!ChangeLayout(kompasDocument))
                {
                    Log.Add(item);
                }
                
            }
            application.Quit();
            if (Log.Count == 0)
            {
                tb_finish.Text = "Готово";
                return;
            }
            using (StreamWriter sw = new StreamWriter("Log.txt",false))
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
    }
}
