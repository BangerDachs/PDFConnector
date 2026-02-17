using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PDFConnector
{

    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<string> _pdfFiles = new ObservableCollection<string>(); // Liste der ausgewählten PDF-Dateien
        private readonly ListBox _pdfFilesList; // ListBox zur Anzeige der PDF-Dateien
        private readonly TextBox _outputPathTextBox; // TextBox zur Anzeige des Zielpfads für die zusammengeführte PDF-Datei

        public MainWindow()
        {
            InitializeComponent();
            _pdfFilesList = FindName("PdfFilesList") as ListBox;
            _outputPathTextBox = FindName("OutputPathTextBox") as TextBox;
            if (_pdfFilesList != null)
            {
                _pdfFilesList.ItemsSource = _pdfFiles;
            }
        }

        private void AddFiles_Click(object sender, RoutedEventArgs e) // "Dateien hinzufügen" Button
        {
            var dialog = new OpenFileDialog
            {
                Filter = "PDF-Dateien (*.pdf)|*.pdf",
                Multiselect = true
            };

            if (dialog.ShowDialog() != true) // Dialog abgebrochen
            {
                return;
            }

            foreach (var file in dialog.FileNames.Where(File.Exists)) // Nur dateie hinzufügen, die existieren
            {
                if (!_pdfFiles.Contains(file))
                {
                    _pdfFiles.Add(file);
                }
            }
        }

        private void RemoveFiles_Click(object sender, RoutedEventArgs e) // "Dateien entfernen" Button
        {
            if (_pdfFilesList == null)
            {
                return;
            }

            var selected = _pdfFilesList.SelectedItems.Cast<string>().ToList();
            foreach (var file in selected) // Ausgewählte Dateien entfernen
            {
                _pdfFiles.Remove(file);
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e) // "Nach oben" Button| Dateien nach oben verschieben
        {
            if (_pdfFilesList == null)
            {
                return;
            }

            var selectedIndex = _pdfFilesList.SelectedIndex; 
            if (selectedIndex <= 0) // Kein Element oder erstes ausgewählt
            {
                return;
            }

            var item = _pdfFiles[selectedIndex];
            _pdfFiles.RemoveAt(selectedIndex);
            _pdfFiles.Insert(selectedIndex - 1, item);
            _pdfFilesList.SelectedIndex = selectedIndex - 1;
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e) // "Nach unten" Button| Dateien nach unten verschieben
        {
            if (_pdfFilesList == null)
            {
                return;
            }

            var selectedIndex = _pdfFilesList.SelectedIndex;
            if (selectedIndex < 0 || selectedIndex >= _pdfFiles.Count - 1) // Kein Element oder letztes ausgewählt
            {
                return;
            }

            var item = _pdfFiles[selectedIndex];
            _pdfFiles.RemoveAt(selectedIndex);
            _pdfFiles.Insert(selectedIndex + 1, item);
            _pdfFilesList.SelectedIndex = selectedIndex + 1;
        }

        private void BrowseOutput_Click(object sender, RoutedEventArgs e) // "Durchsuchen" Button
        {
            var dialog = new SaveFileDialog
            {
                Filter = "PDF-Dateien (*.pdf)|*.pdf",
                FileName = "merged.pdf"
            };

            if (dialog.ShowDialog() == true) // Dialog bestätigt
            {
                if (_outputPathTextBox != null) // Zielpfad im Textfeld anzeigen
                {
                    _outputPathTextBox.Text = dialog.FileName;
                }
            }
        }

        private void Merge_Click(object sender, RoutedEventArgs e) // "Zusammenführen" Button
        {
            if (_pdfFiles.Count == 0) // Keine dateien ausgewählt
            {
                MessageBox.Show("Bitte wählen Sie mindestens eine PDF-Datei aus.", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var outputPath = _outputPathTextBox != null ? _outputPathTextBox.Text : null; // Zielpfad aus dem Textfeld
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                MessageBox.Show("Bitte wählen Sie eine Ziel-Datei aus.", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try // Zusammenführen der PDF dateien
            {
                var outputDirectory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrWhiteSpace(outputDirectory) && !Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                using (var outputDocument = new PdfDocument()) // Neues PDF-Dokument erstellen
                {
                    foreach (var file in _pdfFiles)
                    {
                        using (var inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                        {
                            for (var i = 0; i < inputDocument.PageCount; i++)
                            {
                                outputDocument.AddPage(inputDocument.Pages[i]);
                            }
                        }
                    }

                    outputDocument.Save(outputPath); // Zusammengeführtes PDF speichern
                }

                MessageBox.Show("PDFs wurden zusammengefügt.", "Fertig", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Zusammenfügen: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
