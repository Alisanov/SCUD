using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SCUD
{
    /// <summary>
    /// Логика взаимодействия для MyMessageBox.xaml
    /// </summary>

    public partial class MyMessageBox : Window
    {
        public MyMessageBox()
        {
            InitializeComponent();
            Topmost = true;
        }

        private void AddButtons(MessageBoxButton buttons)
        {
            switch (buttons)
            {
                case MessageBoxButton.OK:
                    AddButton("OK", MessageBoxResult.OK);
                    break;
                case MessageBoxButton.OKCancel:
                    AddButton("OK", MessageBoxResult.OK);
                    AddButton("Отмена", MessageBoxResult.Cancel, isCancel: true);
                    break;
                case MessageBoxButton.YesNo:
                    AddButton("Да", MessageBoxResult.Yes);
                    AddButton("Нет", MessageBoxResult.No);
                    break;
                case MessageBoxButton.YesNoCancel:
                    AddButton("Да", MessageBoxResult.Yes);
                    AddButton("Нет", MessageBoxResult.No);
                    AddButton("Отмена", MessageBoxResult.Cancel, isCancel: true);
                    break;
                default:
                    throw new ArgumentException("Unknown button value", "buttons");
            }
        }
        private void AddButton(string text, MessageBoxResult result, bool isCancel = false)
        {
            Button button = new Button() { Content = text, IsCancel = isCancel, Style = (Style)Application.Current.FindResource("ShadowButton"), Margin = new Thickness(2), Width = 100 };
            button.Click += (o, args) =>
            {
                Result = result;
                DialogResult = true;
            };
            ButtonContainer.Children.Add(button);
        }

        private MessageBoxResult Result = MessageBoxResult.None;

        public static MessageBoxResult Show(string message)
        {
            MyMessageBox dialog = new MyMessageBox();
            dialog.MessageContainer.Text = message;
            dialog.ShowDialog();
            return dialog.Result;
        }
        public static MessageBoxResult Show(string message, string caption)
        {
            MyMessageBox dialog = new MyMessageBox
            {
                Title = caption
            };
            dialog.Caption.Text = caption;
            dialog.MessageContainer.Text = message;
            dialog.ShowDialog();
            return dialog.Result;
        }
        public static MessageBoxResult Show(string message, string caption, MessageBoxButton buttons)
        {
            MyMessageBox dialog = new MyMessageBox
            {
                Title = caption
            };
            dialog.Caption.Text = caption;
            dialog.MessageContainer.Text = message;
            dialog.AddButtons(buttons);
            dialog.ShowDialog();
            return dialog.Result;
        }

        private void MessageBox_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
    }
}

