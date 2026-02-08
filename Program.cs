using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAutoSlider
{
    internal class Program
    {
        // Переменные для управления PowerPoint
        private static dynamic _pptApp;
        private static dynamic _presentation;
        private static System.Windows.Forms.Timer _timer;
        private static NotifyIcon _trayIcon;
        private static ContextMenuStrip _contextMenu;
        private static readonly object _lockObject = new object();

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string filePath = null;

            // 1. Проверяем аргументы запуска (перетаскивание файла на exe)
            if (args.Length > 0)
            {
                filePath = args[0];
            }
            else
            {
                // 2. Если аргументов нет, показываем диалог выбора
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "PowerPoint презентации (*.ppt; *.pptx; *.pps; *.ppsx)|*.ppt;*.pptx;*.pps;*.ppsx";
                    openFileDialog.Title = "Выберите файл для автопоказа";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        filePath = openFileDialog.FileName;
                    }
                    else
                    {
                        // Пользователь отменил выбор, просто выходим
                        return;
                    }
                }
            }

            if (!File.Exists(filePath))
            {
                MessageBox.Show($"Файл не найден:\n{filePath}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 3. Создаем иконку в трее (обязательно ДО запуска презентации)
            CreateTrayIcon(filePath);

            try
            {
                // 4. Запускаем PowerPoint
                StartAutoSlideShow(filePath);

                // 5. Запускаем цикл сообщений Windows
                Application.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Критическая ошибка:\n{ex.Message}", "PowerPoint AutoSlider", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cleanup();
            }
        }

        static void CreateTrayIcon(string filePath)
        {
            _contextMenu = new ContextMenuStrip();
            _contextMenu.Items.Add("О программе", null, (s, e) => MessageBox.Show("PowerPoint AutoSlider v2.0\nminifun.ru", "О программе"));
            _contextMenu.Items.Add("-");
            _contextMenu.Items.Add("Выход", null, OnExit);

            _trayIcon = new NotifyIcon();
            _trayIcon.ContextMenuStrip = _contextMenu;

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                // Убедитесь, что имя ресурса совпадает с вашим Namespace и папкой
                using (var stream = assembly.GetManifestResourceStream("PowerPointAutoSlider.Resources.icon.ico"))
                {
                    if (stream != null)
                    {
                        _trayIcon.Icon = new Icon(stream);
                    }
                    else
                    {
                        _trayIcon.Icon = SystemIcons.Application;
                    }
                }
            }
            catch
            {
                _trayIcon.Icon = SystemIcons.Application;
            }

            string fileName = Path.GetFileName(filePath);
            string title = "AutoSlider: " + fileName;
            if (title.Length > 63) title = title.Substring(0, 60) + "...";
            _trayIcon.Text = title;

            _trayIcon.Visible = true;

            _trayIcon.BalloonTipTitle = "Презентация запущена";
            _trayIcon.BalloonTipText = $"Файл: {fileName}\nИнтервал: 5 сек.";
            _trayIcon.BalloonTipIcon = ToolTipIcon.Info;
            _trayIcon.ShowBalloonTip(3000);
        }

        static void StartAutoSlideShow(string filePath)
        {
            lock (_lockObject)
            {
                if (_pptApp != null) Cleanup();

                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null) throw new Exception("PowerPoint не установлен.");

                _pptApp = Activator.CreateInstance(pptType);
                _pptApp.Visible = -1;

                _presentation = _pptApp.Presentations.Open(filePath, WithWindow: -1);
                _presentation.SlideShowSettings.Run();

                // Ждем появления окна показа слайдов
                int retries = 50;
                while (_pptApp.SlideShowWindows.Count == 0 && retries > 0)
                {
                    Thread.Sleep(200); // Теперь Thread будет найден
                    retries--;
                }

                // Даем время на отрисовку первого слайда
                Thread.Sleep(1500);

                _timer = new System.Windows.Forms.Timer();
                _timer.Interval = 5000;
                _timer.Tick += AutoNextSlide;
                _timer.Start();
            }
        }

        static void AutoNextSlide(object sender, EventArgs e)
        {
            try
            {
                if (_pptApp == null) return;

                if (_pptApp.SlideShowWindows.Count > 0)
                {
                    _pptApp.SlideShowWindows[1].View.Next();
                }
                else
                {
                    _timer?.Stop();
                    _trayIcon.ShowBalloonTip(2000, "Готово", "Показ презентации завершен.", ToolTipIcon.Info);
                }
            }
            catch
            {
                // Игнорируем ошибки
            }
        }

        static void OnExit(object sender, EventArgs e)
        {
            Cleanup();
            Application.Exit();
        }

        static void Cleanup()
        {
            try
            {
                _timer?.Stop();
                _timer?.Dispose();
                _timer = null;

                if (_presentation != null)
                {
                    try { _presentation.Close(); } catch { }
                    Marshal.ReleaseComObject(_presentation);
                    _presentation = null;
                }

                if (_pptApp != null)
                {
                    try { _pptApp.Quit(); } catch { }
                    Marshal.ReleaseComObject(_pptApp);
                    _pptApp = null;
                }

                if (_trayIcon != null)
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                    _trayIcon = null;
                }

                if (_contextMenu != null)
                {
                    _contextMenu.Dispose();
                    _contextMenu = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }
    }
}
