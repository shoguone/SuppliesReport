using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SuppliesReport
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string excelDataFilePath = Properties.Settings.Default.ExcelDataFilePath;
        
        private List<List<PetitionGeneral>> petitions;
        private WordProcessor wordProcessor;

        public MainWindow()
        {
            InitializeComponent();

            FiringDate.SelectedDate = DateTime.Today;
            HiringDate.SelectedDate = DateTime.Today;

            InitializeEnvironment();
            //Application.Current.Shutdown();
        }

        private void FillEverything()
        {
            int i = 0;
            foreach (List<PetitionGeneral> pts in petitions)
            {
                i++;
                wordProcessor.FillActInspection(pts, i);
                wordProcessor.FillActWritingOff(pts, i);

                int j = 0;
                foreach (PetitionGeneral p in pts)
                {
                    j++;
                    wordProcessor.FillActElimination(p, i, j);
                }

                wordProcessor.FillProtocolSession(pts, i);
                wordProcessor.FillOrder(pts, i);
            }
        }

        private void InitializeEnvironment()
        {
            //var excelProcessor = new ExcelProcessor();
            //petitions = excelProcessor.Boo(excelDataFilePath);

            wordProcessor = new WordProcessor();

        }

        private void SupplyPopButton_Click(object sender, RoutedEventArgs e)
        {
            if (FiringDate.SelectedDate.HasValue)
            {
                wordProcessor.FillSupplyPop(FiringDate.SelectedDate.Value);
            }
            else
            {
                MessageBox.Show("Выберите дату снятия с довольствия");
            }
        }

        private void SupplyPushButton_Click(object sender, RoutedEventArgs e)
        {
            if (HiringDate.SelectedDate.HasValue)
            {
                wordProcessor.FillSupplyPush(HiringDate.SelectedDate.Value);
            }
            else
            {
                MessageBox.Show("Выберите дату постановки на довольствие");
            }
        }

        private void ProtocolsButton_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (List<PetitionGeneral> pts in petitions)
            {
                i++;
                wordProcessor.FillProtocolSession(pts, i);
            }
        }

        private void ActsInspectionButton_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (List<PetitionGeneral> pts in petitions)
            {
                i++;
                wordProcessor.FillActInspection(pts, i);
            }
        }

        private void ActsWritingoffButton_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (List<PetitionGeneral> pts in petitions)
            {
                i++;
                wordProcessor.FillActWritingOff(pts, i);
            }
        }

        private void OrdersButton_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (List<PetitionGeneral> pts in petitions)
            {
                i++;
                wordProcessor.FillOrder(pts, i);
            }
        }

        private void ActsEliminationButton_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (List<PetitionGeneral> pts in petitions)
            {
                i++;
                int j = 0;
                foreach (PetitionGeneral p in pts)
                {
                    j++;
                    wordProcessor.FillActElimination(p, i, j);
                }
            }
        }

        private void AllButton_Click(object sender, RoutedEventArgs e)
        {
            FillEverything();
        }





    }
}
