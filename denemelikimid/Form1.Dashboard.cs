using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using denemelikimid.DataBase;

namespace denemelikimid
{
    public partial class Form1
    {
        private sealed class DashboardStats
        {
            public int Toplam { get; }
            public int Aktif { get; }
            public int Puantaj { get; }
            public decimal Odeme { get; }

            public DashboardStats(int toplam, int aktif, int puantaj, decimal odeme)
            {
                Toplam = toplam;
                Aktif = aktif;
                Puantaj = puantaj;
                Odeme = odeme;
            }
        }

        private void LoadDashboardView()
        {
            // 1. SAYFAYI SIFIRLA
            panelContent.Controls.Clear();

            // 2. ANA İSKELET (DİKEY TABLO)
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = colorContent,
                Padding = new Padding(10)
            };

            // Satır Ayarları
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));          // Başlık
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));               // İstatistikler
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));          // Alt kısım

            panelContent.Controls.Add(mainLayout);

            // 3. BAŞLIK ve YENİLE BUTONU
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill };

            Label lblHeader = new Label
            {
                Text = "📊 Genel Durum Paneli",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Location = new Point(0, 10),
                AutoSize = true
            };
            pnlHeader.Controls.Add(lblHeader);

            Button btnRefresh = new Button
            {
                Text = "🔄 Yenile",
                Size = new Size(100, 35),
                BackColor = colorPrimary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Dock = DockStyle.Right
            };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.Resize += (s, e) =>
            {
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();

                int radius = 20;
                path.AddArc(0,0, radius, radius, 180, 90);
                path.AddArc(btnRefresh.Width - radius, 0, radius, radius, 270, 90);
                path.AddArc(btnRefresh.Width - radius, btnRefresh.Height - radius, radius, radius, 0, 90);
                path.AddArc(0, btnRefresh.Height - radius, radius, radius, 90, 90);

                path.CloseFigure();
                btnRefresh.Region = new Region(path);
            };
            btnRefresh.Click += (s, e) => LoadDashboardView();
            pnlHeader.Controls.Add(btnRefresh);

            mainLayout.Controls.Add(pnlHeader, 0, 0);

            // 4. İSTATİSTİK KARTLARI
            FlowLayoutPanel pnlStats = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Margin = new Padding(0, 0, 0, 20)
            };

            Label lblToplam;
            Label lblAktif;
            Label lblPuantaj;
            Label lblOdeme;

            pnlStats.Controls.Add(CreateStatCard("👥 Toplam Personel", "...", colorPrimary, out lblToplam));
            pnlStats.Controls.Add(CreateStatCard("✅ Aktif Çalışan", "...", colorSuccess, out lblAktif));
            pnlStats.Controls.Add(CreateStatCard("📝 Bu Ay Puantaj", "...", Color.Orange, out lblPuantaj));
            pnlStats.Controls.Add(CreateStatCard("💰 Toplam Ödeme", "...", colorInfo, out lblOdeme));

            mainLayout.Controls.Add(pnlStats, 0, 1);

            // 5. ALT KISIM (BUTONLAR ve LOG)
            TableLayoutPanel bottomLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent
            };
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            // Sol: Modern Takvim + Hatırlatmalar
            Panel pnlLeft = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10),
                Margin = new Padding(0, 0, 10, 0)
            };
            Label lblLeft = new Label
            {
                Text = "📅 Takvim ve Hatırlatmalar",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 30
            };

            Action<Control, int> applyRoundedCard = (ctrl, radius) =>
            {
                Action refreshRegion = () =>
                {
                    if (ctrl.Width <= 0 || ctrl.Height <= 0) return;
                    ctrl.Region = new Region(CreateRoundedRectPath(new Rectangle(0, 0, ctrl.Width - 1, ctrl.Height - 1), radius));
                };

                ctrl.SizeChanged += (s, e) => refreshRegion();
                refreshRegion();
            };

            TableLayoutPanel leftLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 72F));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 28F));

            Panel pnlCalendarCard = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(8)
            };

            Panel pnlCalendarTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 34,
                BackColor = Color.FromArgb(245, 248, 255)
            };

            Label lblCalendarTop = new Label
            {
                Text = DateTime.Today.ToString("MMMM yyyy", new System.Globalization.CultureInfo("tr-TR")),
                Dock = DockStyle.Left,
                Width = 170,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                TextAlign = ContentAlignment.MiddleLeft
            };

            Button btnToday = new Button
            {
                Text = "Bugün",
                Dock = DockStyle.Right,
                Width = 70,
                BackColor = colorPrimary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
            };
            btnToday.FlatAppearance.BorderSize = 0;
            ApplyRoundedCorners(btnToday, 8);

            pnlCalendarTop.Controls.Add(btnToday);
            pnlCalendarTop.Controls.Add(lblCalendarTop);

            Panel pnlCalendarBody = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 6, 0, 0)
            };

            MonthCalendar miniCalendar = new MonthCalendar
            {
                Dock = DockStyle.Fill,
                MaxSelectionCount = 1,
                BackColor = Color.White,
                TitleBackColor = Color.FromArgb(67, 97, 238),
                TitleForeColor = Color.White,
                TrailingForeColor = Color.FromArgb(160, 160, 160)
            };
            pnlCalendarBody.Controls.Add(miniCalendar);
            pnlCalendarCard.Controls.Add(pnlCalendarBody);
            pnlCalendarCard.Controls.Add(pnlCalendarTop);

            Panel pnlReminder = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 4, 0, 0)
            };

            Label lblReminder = new Label
            {
                Text = "🔔 Seçili Gün Hatırlatmaları",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 22
            };

            Button btnOpenReminderDialog = new Button
            {
                Text = "+ Hatırlatma Ekle",
                Dock = DockStyle.Top,
                Height = 34,
                BackColor = colorPrimary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 6)
            };
            btnOpenReminderDialog.FlatAppearance.BorderSize = 0;
            ApplyRoundedCorners(btnOpenReminderDialog, 10);

            ListBox lstReminder = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9.5f),
                BackColor = Color.WhiteSmoke,
                ItemHeight = 18
            };

            Action loadReminderList = () =>
            {
                try
                {
                    EnsureDashboardReminderTable();
                    LoadDashboardReminders(lstReminder, miniCalendar.SelectionStart.Date);
                    lblCalendarTop.Text = miniCalendar.SelectionStart.ToString("MMMM yyyy", new System.Globalization.CultureInfo("tr-TR"));
                }
                catch
                {
                    // DB erişimi yoksa ekran çalışmaya devam etsin
                }
            };

            btnToday.Click += (s, e) =>
            {
                miniCalendar.SetDate(DateTime.Today);
                loadReminderList();
            };

            btnOpenReminderDialog.Click += (s, e) =>
            {
                using (Form dlg = new Form())
                {
                    dlg.Text = "Hatırlatma Ekle";
                    dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                    dlg.StartPosition = FormStartPosition.CenterParent;
                    dlg.MaximizeBox = false;
                    dlg.MinimizeBox = false;
                    dlg.ClientSize = new Size(420, 190);
                    dlg.BackColor = colorContent;

                    Label lblDate = new Label
                    {
                        Text = "Tarih: " + miniCalendar.SelectionStart.ToString("dd.MM.yyyy"),
                        AutoSize = true,
                        Font = new Font("Segoe UI", 10, FontStyle.Bold),
                        Location = new Point(15, 15)
                    };

                    DateTimePicker dtpTime = new DateTimePicker
                    {
                        Format = DateTimePickerFormat.Custom,
                        CustomFormat = "HH:mm",
                        ShowUpDown = true,
                        Font = new Font("Segoe UI", 10),
                        Width = 90,
                        Location = new Point(15, 45)
                    };

                    TextBox txtReminder = new TextBox
                    {
                        Font = new Font("Segoe UI", 10),
                        Width = 380,
                        Height = 30,
                        Location = new Point(15, 80)
                    };

                    Button btnSave = new Button
                    {
                        Text = "Kaydet",
                        Width = 95,
                        Height = 34,
                        Location = new Point(300, 130),
                        BackColor = colorPrimary,
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat
                    };
                    btnSave.FlatAppearance.BorderSize = 0;
                    ApplyRoundedCorners(btnSave, 10);

                    Button btnCancel = new Button
                    {
                        Text = "İptal",
                        Width = 95,
                        Height = 34,
                        Location = new Point(195, 130),
                        BackColor = Color.FromArgb(148, 163, 184),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat
                    };
                    btnCancel.FlatAppearance.BorderSize = 0;
                    ApplyRoundedCorners(btnCancel, 10);

                    btnSave.Click += (sx, ex) =>
                    {
                        string text = txtReminder.Text.Trim();
                        if (string.IsNullOrEmpty(text))
                        {
                            MessageBox.Show("Hatırlatma metni boş olamaz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        try
                        {
                            AddDashboardReminder(miniCalendar.SelectionStart.Date, dtpTime.Value.TimeOfDay, text);
                            dlg.DialogResult = DialogResult.OK;
                            dlg.Close();
                        }
                        catch
                        {
                            MessageBox.Show("Hatırlatma eklenemedi.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    };

                    btnCancel.Click += (sx, ex) => dlg.Close();

                    txtReminder.KeyDown += (sx, ex) =>
                    {
                        if (ex.KeyCode != Keys.Enter) return;
                        btnSave.PerformClick();
                        ex.SuppressKeyPress = true;
                    };

                    dlg.Controls.Add(lblDate);
                    dlg.Controls.Add(dtpTime);
                    dlg.Controls.Add(txtReminder);
                    dlg.Controls.Add(btnCancel);
                    dlg.Controls.Add(btnSave);

                    if (dlg.ShowDialog(this) == DialogResult.OK)
                        loadReminderList();
                }
            };

            miniCalendar.DateChanged += (s, e) => loadReminderList();

            lstReminder.DoubleClick += (s, e) =>
            {
                if (lstReminder.SelectedItem == null) return;
                var selected = lstReminder.SelectedItem as DashboardReminderItem;
                if (selected == null) return;

                var dr = MessageBox.Show("Seçili hatırlatma silinsin mi?", "Hatırlatma Sil", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;

                try
                {
                    DeleteDashboardReminder(selected.Id);
                    loadReminderList();
                }
                catch { }
            };

            loadReminderList();

            pnlReminder.Controls.Add(lstReminder);
            pnlReminder.Controls.Add(btnOpenReminderDialog);
            pnlReminder.Controls.Add(lblReminder);

            leftLayout.Controls.Add(pnlCalendarCard, 0, 0);
            leftLayout.Controls.Add(pnlReminder, 0, 1);

            pnlLeft.Controls.Add(leftLayout);
            pnlLeft.Controls.Add(lblLeft);
            applyRoundedCard(pnlLeft, 14);
            applyRoundedCard(pnlCalendarCard, 12);
            applyRoundedCard(pnlReminder, 12);
            bottomLayout.Controls.Add(pnlLeft, 0, 0);

            // Sağ: Sistem Logu
            Panel pnlRight = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10)
            };
            Label lblRight = new Label
            {
                Text = "📢 Sistem Durumu",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 30
            };
            ListBox lstLog = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 10),
                BackColor = Color.WhiteSmoke,
                ItemHeight = 20
            };

            lstLog.Items.Add($"> [{DateTime.Now:HH:mm}] Panel yüklendi.");
            lstLog.Items.Add($"> Veriler yükleniyor...");

            pnlRight.Controls.Add(lstLog);
            pnlRight.Controls.Add(lblRight);
            bottomLayout.Controls.Add(pnlRight, 1, 0);

            mainLayout.Controls.Add(bottomLayout, 0, 2);

            Task.Run(() => FetchDashboardStats())
                .ContinueWith(t =>
                {
                    if (IsDisposed || panelContent.IsDisposed)
                    {
                        return;
                    }

                    if (t.IsFaulted)
                    {
                        lstLog.Items.Add("> Veriler yüklenemedi.");
                        return;
                    }

                    var stats = t.Result;
                    lblToplam.Text = stats.Toplam.ToString();
                    lblAktif.Text = stats.Aktif.ToString();
                    lblPuantaj.Text = stats.Puantaj.ToString();
                    lblOdeme.Text = stats.Odeme.ToString("C0");

                    lstLog.Items.Clear();
                    lstLog.Items.Add($"> [{DateTime.Now:HH:mm}] Panel yüklendi.");
                    lstLog.Items.Add("> Veritabanı bağlantısı: OK");
                    lstLog.Items.Add($"> Toplam {stats.Toplam} personel mevcut.");
                    if (stats.Puantaj == 0)
                        lstLog.Items.Add("> UYARI: Bu ay henüz puantaj girilmemiş!");
                    else
                        lstLog.Items.Add($"> {stats.Puantaj} personelin puantajı hazır.");
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private DashboardStats FetchDashboardStats()
        {
            int toplam = 0;
            int aktif = 0;
            int puantaj = 0;
            decimal odeme = 0;

            using (var conn = DbConnection.GetConnection())
            {
                conn.Open();
                using (var cmd = new SqliteCommand("SELECT COUNT(*) FROM program_katilimcilari", conn))
                    toplam = Convert.ToInt32(cmd.ExecuteScalar());
                using (var cmd = new SqliteCommand("SELECT COUNT(*) FROM program_katilimcilari WHERE pk_isten_ayrilma_tarihi IS NULL", conn))
                    aktif = Convert.ToInt32(cmd.ExecuteScalar());

                string buAy = DateTime.Now.ToString("yyyy-MM");
                using (var cmd = new SqliteCommand("SELECT COUNT(*) FROM puantaj WHERE p_yil_ay = @ay AND p_calistigi_gun_sayisi > 0", conn))
                {
                    cmd.Parameters.AddWithValue("@ay", buAy);
                    puantaj = Convert.ToInt32(cmd.ExecuteScalar());
                }

                using (var cmd = new SqliteCommand("SELECT SUM(b_odenmesi_gereken_net_tutar) FROM bordro", conn))
                {
                    var res = cmd.ExecuteScalar();
                    if (res != DBNull.Value) odeme = Convert.ToDecimal(res);
                }
            }

            return new DashboardStats(toplam, aktif, puantaj, odeme);
        }

        private void EnsureDashboardReminderTable()
        {
            using (var conn = DbConnection.GetConnection())
            {
                conn.Open();
                string sql = @"CREATE TABLE IF NOT EXISTS dashboard_reminders (
                                dr_id INTEGER PRIMARY KEY AUTOINCREMENT,
                                dr_date TEXT NOT NULL,
                                dr_time TEXT NOT NULL,
                                dr_text TEXT NOT NULL,
                                dr_created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                               );";
                new SqliteCommand(sql, conn).ExecuteNonQuery();
            }
        }

        private void LoadDashboardReminders(ListBox listBox, DateTime selectedDate)
        {
            listBox.Items.Clear();

            using (var conn = DbConnection.GetConnection())
            {
                conn.Open();

                var cmd = new SqliteCommand(@"SELECT dr_id, dr_time, dr_text 
                                             FROM dashboard_reminders 
                                             WHERE dr_date = @date 
                                             ORDER BY dr_time ASC", conn);
                cmd.Parameters.AddWithValue("@date", selectedDate.Date);

                using (var dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        listBox.Items.Add(new DashboardReminderItem
                        {
                            Id = Convert.ToInt32(dr["dr_id"]),
                            Time = dr["dr_time"].ToString(),
                            Text = dr["dr_text"].ToString()
                        });
                    }
                }
            }
        }

        private void AddDashboardReminder(DateTime date, TimeSpan time, string text)
        {
            using (var conn = DbConnection.GetConnection())
            {
                conn.Open();
                var cmd = new SqliteCommand("INSERT INTO dashboard_reminders (dr_date, dr_time, dr_text) VALUES (@date, @time, @text)", conn);
                cmd.Parameters.AddWithValue("@date", date.Date);
                cmd.Parameters.AddWithValue("@time", time);
                cmd.Parameters.AddWithValue("@text", text);
                cmd.ExecuteNonQuery();
            }
        }

        private void DeleteDashboardReminder(int id)
        {
            using (var conn = DbConnection.GetConnection())
            {
                conn.Open();
                var cmd = new SqliteCommand("DELETE FROM dashboard_reminders WHERE dr_id = @id", conn);
                cmd.Parameters.AddWithValue("@id", id);
                cmd.ExecuteNonQuery();
            }
        }

        private class DashboardReminderItem
        {
            public int Id { get; set; }
            public string Time { get; set; }
            public string Text { get; set; }
            public override string ToString() => $"[{Time}] {Text}";
        }

        private Button CreateQuickBtn(string text, Color color, Action onClickAction)
        {
            Button btn = new Button
            {
                Text = text,
                Size = new Size(200, 100),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 20, 20)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) => onClickAction.Invoke();
            ApplyRoundedCorners(btn, 15);
            return btn;
        }

        private Panel CreateStatCard(string title, string value, Color color, out Label valueLabel)
        {
            Panel card = new Panel
            {
                Width = 250,
                Height = 140,
                BackColor = Color.White,
                Margin = new Padding(0, 0, 20, 0)
            };

            Panel accent = new Panel { Dock = DockStyle.Left, Width = 5, BackColor = color };
            card.Controls.Add(accent);

            Label lblValue = new Label
            {
                Text = value,
                Font = new Font("Segoe UI", 24, FontStyle.Bold),
                ForeColor = Color.Black,
                Location = new Point(20, 25),
                AutoSize = true
            };
            card.Controls.Add(lblValue);
            valueLabel = lblValue;

            Label lblTitle = new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 11, FontStyle.Regular),
                ForeColor = Color.Gray,
                Location = new Point(20, 80),
                AutoSize = true
            };
            card.Controls.Add(lblTitle);

            return card;
        }

        private DataGridView CreateMiniGrid()
        {
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            return dgv;
        }
    }
}




