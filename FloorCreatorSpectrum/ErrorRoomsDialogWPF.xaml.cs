using Autodesk.Revit.DB.Architecture;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace FloorCreator
{
    public partial class ErrorRoomsDialogWPF : Window
    {
        private List<Room> _errorRooms;

        public ErrorRoomsDialogWPF(List<Room> errorRooms)
        {
            InitializeComponent();

            _errorRooms = errorRooms ?? new List<Room>();

            var roomDescriptions = new List<string>(_errorRooms.Count);

            foreach (Room room in _errorRooms)
            {
                if (room == null) continue;

#if R2019 || R2020 || R2021 || R2022 || R2023 || R2024 || R2025
                // Revit 2019–2025
                int idValue = room.Id.IntegerValue;
                roomDescriptions.Add($"ID: {idValue}, Номер: {room.Number}, Имя: {room.Name}");
#else
                // Revit 2026+
                long idValue = room.Id.Value;
                roomDescriptions.Add($"ID: {idValue}, Номер: {room.Number}, Имя: {room.Name}");
#endif
            }

            lstErrorRooms.ItemsSource = roomDescriptions;
        }


        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                FileName = "ErrorRooms.txt",
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                using (StreamWriter writer = new StreamWriter(saveFileDialog.FileName))
                {
                    writer.WriteLine("Не удалось создать полы в следующих помещениях:");

                    foreach (Room room in _errorRooms)
                    {
                        if (room == null) continue;

#if R2019 || R2020 || R2021 || R2022 || R2023 || R2024 || R2025
                        writer.WriteLine($"ID: {room.Id.IntegerValue}, Номер: {room.Number}, Имя: {room.Name}");
#else
                        writer.WriteLine($"ID: {room.Id.Value}, Номер: {room.Number}, Имя: {room.Name}");
#endif
                    }
                }

                MessageBox.Show(
                    $"Список помещений сохранен в {saveFileDialog.FileName}",
                    "Сохранение завершено",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }
    }
}
