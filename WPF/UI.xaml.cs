using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using System.IO;

namespace BaseOffice.UI
{
    /// <summary>
    /// Interaction logic for UI.xaml
    /// </summary>
    public partial class UI : Window
    {
        public string City { get; set; }
        public string State { get; set; }

        private Document Doc;

        

        public UI(Document doc)
        {
            InitializeComponent();
            Doc = doc;

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            var iconPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(thisAssemblyPath), "Resources", "closedButton.png");

            this.ClosedButtonImage.Source = new BitmapImage(new Uri(iconPath));
        }

        private void Image_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.Close();
        }

        private void ButtonOK_Click(object sender, RoutedEventArgs e)
        {
            City = this.CityBox.Text;
            State = this.StateBox.Text;

            this.Close();

        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Verifica se o botão esquerdo do mouse está pressionado
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                // Permite arrastar a janela
                this.DragMove();
            }
        }
    }
}
