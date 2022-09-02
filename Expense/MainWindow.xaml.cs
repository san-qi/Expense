using Microsoft.Data.Sqlite;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Expense
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        private Grid Main_content;
        private Grid Last_content;
        public MainWindow()
        {
            InitializeComponent();

            Recption_reimbursement_date.Text = DateTime.Now.ToString("yyyy.MM.dd");
            Recption_tax_count.Text = "0";
            Recption_tax_paper_number.Text = "1";
            using (var conn = new SqliteConnection("Data Source=conf.db"))
            {
                conn.Open();
                var command = conn.CreateCommand();
                command.CommandText = @"Create table if not exists config (name varchar(20), source varchar(100))";
                command.ExecuteNonQuery();
                command.CommandText = @"Create table if not exists config (name varchar(20), score varchar(100))";
                command.ExecuteNonQuery();
                command.CommandText = @"Insert into config (name, source) values ('recption_path', '')";
                command.ExecuteNonQuery();
            }
            Main_content = Recption_content;
            Last_content = Menu_content;
        }

        private void Change_content(Grid content)
        {
            Last_content = Main_content;
            Main_content = content;
            Last_content.Visibility = Visibility.Hidden;
            Main_content.Visibility = Visibility.Visible;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void Date_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var input_box = sender as TextBox;
            string input = input_box.Text + e.Text;
            var re1 = new Regex(@"^\d{0,4}$");
            var re2 = new Regex(@"^\d{0,4}[ .\\/-]\d{0,2}$");
            var re3 = new Regex(@"^\d{0,4}[ .\\/-]\d{0,2}[ .\\/-]\d{0,2}$");
            e.Handled = !re1.IsMatch(input) && !re2.IsMatch(input) && !re3.IsMatch(input);
        }

        private void Digit_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var input_box = sender as TextBox;
            string input = input_box.Text + e.Text;
            var re1 = new Regex(@"^\d{0,1}$");
            e.Handled = !re1.IsMatch(input);
        }

        private void Account_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var input_box = sender as TextBox;
            string input = input_box.Text + e.Text;
            var re1 = new Regex(@"^\d*$");
            var re2 = new Regex(@"^\d*\.\d{0,2}$");
            e.Handled = !re1.IsMatch(input) && !re2.IsMatch(input);
        }

        private void Submit_click(object sender, RoutedEventArgs e)
        {
            if (Recption_name.Name.Length > 0 && Recption_reception_employer.Text.Length > 0 && Recption_reception_people.Text.Length > 0 &&
                        Recption_target_place.Text.Length > 0 && Recption_total_count.Text.Length > 0 && Recption_tax_count.Text.Length > 0 && Recption_tax_paper_number.Text.Length > 0 &&
                        Recption_start_date.Text.Length > 0 && Recption_reimbursement_date.Text.Length > 0 && Recption_reason.Text.Length > 0)
            {
                Expense.Program.Generate_reception(Recption_name.Text, Recption_colleagues.Text, Recption_reception_employer.Text, Recption_reception_people.Text,
                        Recption_target_place.Text, Recption_meal_time.Text, Recption_total_count.Text, Recption_tax_count.Text, Recption_tax_paper_number.Text,
                        Recption_start_date.Text, Recption_reimbursement_date.Text, Recption_reason.Text, Recption_have_wine_paper.Text == "是");
                Mask.Content = "报销单生成成功！";
            }
            else
            {
                Mask.Content = "请将信息补全完整！";
            }
            Change_content(Mask_content);
        }
        private void Mask_click(object sender, RoutedEventArgs e)
        {
            Change_content(Last_content);
        }

        private void Quit_click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void Menu_click(object sender, RoutedEventArgs e)
        {
            if (Menu_content.Visibility == Visibility.Hidden)
            {
                Change_content(Menu_content);
            }
            else
            {
                Change_content(Last_content);
            }
        }

        private void Menu_back_click(object sender, RoutedEventArgs e)
        {
            Change_content(Last_content);
        }

        private void Menu_chose_recption_path_click(object sender, RoutedEventArgs e)
        {
            using (var conn = new SqliteConnection("Data Source=conf.db"))
            {
                conn.Open();
                var command = conn.CreateCommand();

                var dialog = new CommonOpenFileDialog()
                {
                    IsFolderPicker = true,
                    Title = "请选择生成报销清单所在的文件夹",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                };
                if (dialog.ShowDialog() == CommonFileDialogResult.Cancel)
                {
                    command.CommandText = @"Update config set source = '' where name = 'recption_path'";
                    command.ExecuteNonQuery();
                }
                else
                {
                    command.CommandText = String.Format("Update config set source = '{0}' where name = 'recption_path'", dialog.FileName);
                    command.ExecuteNonQuery();

                    Main_content.Visibility = Visibility.Hidden;
                    Main_content = Last_content;
                    Mask.Content = "修改路径成功!";
                    Change_content(Mask_content);
                }
            }
        }

        private void Change_click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var white = new SolidColorBrush(Colors.White);
            var black = new SolidColorBrush(Colors.Black);
            if ((btn.Background as SolidColorBrush).Color == black.Color)
            {
                btn.Background = white;
            }
            else
            {
                btn.Background = black;
            }
        }
    }
}
