using ChangeLayoutStyle.Properties;
using Kompas6API5;
using KompasAPI7;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
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
            #endregion
        }


        private void ChangeLayout(IKompasDocument kompasDocument)
        {
            if (kompasDocument is null)
            {
                tb_finish.Text = "Не получилось открыть документ";
                return;
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
            kompasDocument.Close(Kompas6Constants.DocumentCloseOptions.kdSaveChanges);

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
            tb_finish.Text = "";
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


            string[]  FilesDirs = Directory.GetFiles(tb_folderDir.Text, "*.cdw");
            foreach (string item in FilesDirs)
            {
                ChangeLayout(documets.Open(item, false, false));
            }
            application.Quit();

            tb_finish.Text = "Готово";
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            #region Сохранение настроек
            Settings.Default.FolderDir = tb_folderDir.Text;
            Settings.Default.LayoutLibraryFileName = tb_layoutLibraryFileName.Text;
            Settings.Default.Save();
            #endregion
        }
    }
}
