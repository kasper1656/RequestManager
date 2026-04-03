using System;
using System.Linq;
using System.Windows;
using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Collections.Generic;

namespace RequestManagerApp
{
    public partial class MainWindow : Window
    {
        private User _currentUser;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var db = new AppDbContext())
                {
                    bool canConnect = db.Database.CanConnect();
                    if (!canConnect)
                    {
                        MessageBox.Show("Не удалось подключиться к MySQL!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        Application.Current.Shutdown();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка:\n{ex.Message}", "Сбой", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LinkToRegister_Click(object sender, RoutedEventArgs e) { PanelLogin.Visibility = Visibility.Collapsed; PanelRegister.Visibility = Visibility.Visible; }
        private void LinkToLogin_Click(object sender, RoutedEventArgs e) { PanelRegister.Visibility = Visibility.Collapsed; PanelLogin.Visibility = Visibility.Visible; }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new AppDbContext())
            {
                var user = db.Users.FirstOrDefault(u => u.Username == txtLoginUser.Text.Trim() && u.Password == txtLoginPass.Password.Trim());
                if (user != null) LoginSuccess(user);
                else MessageBox.Show("Неверный логин или пароль!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnRegister_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new AppDbContext())
            {
                if (db.Users.Any(u => u.Username == txtRegLogin.Text.Trim())) return;
                var newUser = new User { Name = txtRegName.Text.Trim(), Username = txtRegLogin.Text.Trim(), Password = txtRegPass.Password.Trim(), Role = "User" };
                db.Users.Add(newUser);
                db.SaveChanges();
                LoginSuccess(newUser);
            }
        }

        private void LoginSuccess(User user)
        {
            _currentUser = user;
            txtUserName.Text = $"{_currentUser.Name}\nРоль: {_currentUser.Role}";
            AuthScreen.Visibility = Visibility.Collapsed;
            MainScreen.Visibility = Visibility.Visible;

            if (_currentUser.Role == "User")
            {
                NavBtnAdmin.Visibility = Visibility.Collapsed;
                colWorkType.Visibility = Visibility.Collapsed;
                btnCompleteReq.Visibility = Visibility.Collapsed;
                lblFilterType.Visibility = Visibility.Collapsed;
                cbFilterType.Visibility = Visibility.Collapsed;
                NavUser_Click(null, null);
            }
            else
            {
                NavBtnAdmin.Visibility = Visibility.Visible;
                colWorkType.Visibility = Visibility.Visible;
                btnCompleteReq.Visibility = Visibility.Visible;
                lblFilterType.Visibility = Visibility.Visible;
                cbFilterType.Visibility = Visibility.Visible;
                NavAdmin_Click(null, null);
            }
            LoadData();
        }

        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            _currentUser = null;
            MainScreen.Visibility = Visibility.Collapsed;
            AuthScreen.Visibility = Visibility.Visible;
        }

        private void NavAdmin_Click(object sender, RoutedEventArgs e) => MainTabControl.SelectedIndex = 0;
        private void NavUser_Click(object sender, RoutedEventArgs e) => MainTabControl.SelectedIndex = 1;
        private void NavArchive_Click(object sender, RoutedEventArgs e) => MainTabControl.SelectedIndex = 2;

        private void LoadData()
        {
            using (var db = new AppDbContext())
            {
                if (_currentUser.Role == "Admin")
                {
                    dgFacilities.ItemsSource = db.Facilities.ToList();
                    dgServices.ItemsSource = db.Services.ToList();
                }
                cbFacilities.ItemsSource = db.Facilities.ToList();
                cbServices.ItemsSource = db.Services.ToList();

                var types = db.Services.Select(s => s.WorkType).Distinct().ToList();
                types.Insert(0, "Все типы");
                cbFilterType.ItemsSource = types;
                cbFilterType.SelectedIndex = 0;

                LoadArchive();
            }
        }

        private void Dg_CellEditEnding(object sender, System.Windows.Controls.DataGridCellEditEndingEventArgs e)
        {
            using (var db = new AppDbContext())
            {
                if (sender == dgFacilities) db.Facilities.Update((Facility)e.Row.Item);
                if (sender == dgServices) db.Services.Update((Service)e.Row.Item);
                db.SaveChanges();
            }
            LoadData();
        }

        private void ImportFacilities_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx" };
            if (ofd.ShowDialog() == true)
            {
                try
                {
                    using (var wb = new XLWorkbook(ofd.FileName))
                    using (var db = new AppDbContext())
                    {
                        var ws = wb.Worksheet(1);
                        var rows = ws.RangeUsed().RowsUsed().Skip(1);
                        foreach (var row in rows)
                        {
                            if (row.Cell(1).IsEmpty()) continue;
                            db.Facilities.Add(new Facility
                            {
                                Name = row.Cell(1).GetString(),
                                Address = row.Cell(2).GetString(),
                                ResponsiblePerson = row.Cell(3).GetString()
                            });
                        }
                        db.SaveChanges();
                        MessageBox.Show("Объекты успешно импортированы!");
                        LoadData();
                    }
                }
                catch (Exception ex) { MessageBox.Show("Ошибка импорта: " + ex.Message); }
            }
        }

        private void ImportServices_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx" };
            if (ofd.ShowDialog() == true)
            {
                try
                {
                    using (var wb = new XLWorkbook(ofd.FileName))
                    using (var db = new AppDbContext())
                    {
                        var ws = wb.Worksheet(1);
                        var rows = ws.RangeUsed().RowsUsed().Skip(1);
                        foreach (var row in rows)
                        {
                            if (row.Cell(1).IsEmpty()) continue;
                            db.Services.Add(new Service
                            {
                                Name = row.Cell(1).GetString(),
                                WorkType = row.Cell(2).GetString(),
                                DeadlineDays = row.Cell(3).GetValue<int>()
                            });
                        }
                        db.SaveChanges();
                        MessageBox.Show("Услуги успешно импортированы!");
                        LoadData();
                    }
                }
                catch (Exception ex) { MessageBox.Show("Ошибка импорта: " + ex.Message); }
            }
        }

        private void CreateRequest_Click(object sender, RoutedEventArgs e)
        {
            var fac = cbFacilities.SelectedItem as Facility;
            var svc = cbServices.SelectedItem as Service;

            if (fac == null || svc == null || !int.TryParse(txtQuantity.Text, out int qty)) { MessageBox.Show("Заполните все поля!"); return; }

            using (var db = new AppDbContext())
            {
                db.Requests.Add(new Request
                {
                    UserId = _currentUser.Id,
                    FacilityId = fac.Id,
                    Quantity = qty,
                    ServiceName = svc.Name,
                    WorkType = svc.WorkType,
                    DeadlineDays = svc.DeadlineDays
                });
                db.SaveChanges();
                MessageBox.Show("Заявка успешно создана!");
                LoadArchive();
            }
        }

        private void Filter_Click(object sender, RoutedEventArgs e) => LoadArchive();
        private void ResetFilter_Click(object sender, RoutedEventArgs e)
        {
            dpStart.SelectedDate = null;
            dpEnd.SelectedDate = null;
            cbFilterType.SelectedIndex = 0;
            chkShowCompleted.IsChecked = false;
            LoadArchive();
        }

        private void LoadArchive()
        {
            using (var db = new AppDbContext())
            {
                var query = db.Requests.Include(r => r.Facility).AsQueryable();

                if (chkShowCompleted.IsChecked != true) query = query.Where(r => r.Status == "В очереди");

                if (_currentUser.Role == "User") query = query.Where(r => r.UserId == _currentUser.Id);

                if (dpStart.SelectedDate.HasValue) query = query.Where(r => r.CreatedAt >= dpStart.SelectedDate.Value);
                if (dpEnd.SelectedDate.HasValue)
                {
                    var endOfDay = dpEnd.SelectedDate.Value.AddDays(1).AddTicks(-1);
                    query = query.Where(r => r.CreatedAt <= endOfDay);
                }

                if (cbFilterType.SelectedItem != null && cbFilterType.SelectedItem.ToString() != "Все типы")
                {
                    string wt = cbFilterType.SelectedItem.ToString();
                    query = query.Where(r => r.WorkType == wt);
                }

                dgRequests.ItemsSource = query.OrderByDescending(r => r.CreatedAt).ToList();
            }
        }

        private void CompleteRequest_Click(object sender, RoutedEventArgs e)
        {
            if (dgRequests.SelectedItem is Request req)
            {
                using (var db = new AppDbContext())
                {
                    db.Requests.Find(req.Id).Status = "Выполнена";
                    db.SaveChanges();
                }
                LoadArchive();
            }
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            var exportData = dgRequests.ItemsSource as List<Request>;
            if (exportData == null || exportData.Count == 0)
            {
                MessageBox.Show("Нет данных для выгрузки. Настройте фильтры!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx", FileName = "Акты_Выполненных_Работ.xlsx" };
            if (sfd.ShowDialog() != true) return;

            try
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Акты");
                    int row = 1;

                    string[] headers = { "№ Заявки", "Объект", "Наименование работы", "Срок(дн)", "Кол-во", "Статус", "Дата создания" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = ws.Cell(row, i + 1);
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }
                    row++;

                    int totalQuantity = 0;

                    foreach (var r in exportData)
                    {
                        ws.Cell(row, 1).Value = r.Id;
                        ws.Cell(row, 2).Value = r.Facility.Name;
                        ws.Cell(row, 3).Value = r.ServiceName;
                        ws.Cell(row, 4).Value = r.DeadlineDays;
                        ws.Cell(row, 5).Value = r.Quantity;
                        ws.Cell(row, 6).Value = r.Status;
                        ws.Cell(row, 7).Value = r.FormattedDate;

                        for (int i = 1; i <= 7; i++)
                        {
                            ws.Cell(row, i).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }

                        totalQuantity += r.Quantity;
                        row++;
                    }

                    var totalTextRange = ws.Range(row, 1, row, 4);
                    totalTextRange.Merge();
                    totalTextRange.Value = "итого";
                    totalTextRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    totalTextRange.Style.Font.Bold = true;
                    totalTextRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    var sumCell = ws.Cell(row, 5);
                    sumCell.Value = totalQuantity;
                    sumCell.Style.Font.Bold = true;
                    sumCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    ws.Cell(row, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(row, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    ws.Columns().AdjustToContents();

                    wb.SaveAs(sfd.FileName);
                    MessageBox.Show("Файл Excel успешно сформирован!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка выгрузки: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}