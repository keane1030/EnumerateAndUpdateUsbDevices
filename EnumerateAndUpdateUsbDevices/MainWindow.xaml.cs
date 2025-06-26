using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Management;
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
using System.Windows.Threading;

namespace EnumerateAndUpdateUsbDevices
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly WqlEventQuery _insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent  WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity'");
        private readonly WqlEventQuery _deleteQuery = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent  WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity'");
        private ManagementEventWatcher _insertWatcher;
        private ManagementEventWatcher _deleteWatcher;
        private Dispatcher _dispatcher;

        public MainWindow()
        {
            _dispatcher = Dispatcher.CurrentDispatcher;
            InitializeComponent();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _insertWatcher = new ManagementEventWatcher(_insertQuery);
            _insertWatcher.EventArrived += new EventArrivedEventHandler(DeviceInsertedEvent);

            _insertWatcher.Start();

            _deleteWatcher = new ManagementEventWatcher(_deleteQuery);
            _deleteWatcher.EventArrived += new EventArrivedEventHandler(DeviceDeletedEvent);

            _deleteWatcher.Start();
        }

        private void DeviceInsertedEvent(object sender, EventArrivedEventArgs e)
        {
            var instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];

            try
            {
                string deviceCaption = instance.GetPropertyValue("Caption")?.ToString();

                _dispatcher.Invoke(() =>
                {
                    DeviceTextBlock.Text = "Device Inserted: " + deviceCaption;
                });
            }
            catch (Exception ex)
            {
                _dispatcher.Invoke(() =>
                {
                    DeviceTextBlock.Text = "Error: " + ex.Message;
                });
            }
        }

        private void DeviceDeletedEvent(object sender, EventArrivedEventArgs e)
        {
            var instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];

            try
            {
                string deviceCaption = instance.GetPropertyValue("Caption")?.ToString();

                _dispatcher.Invoke(() =>
                {
                    DeviceTextBlock.Text = "Device Removed: " + deviceCaption;
                });
            }
            catch (Exception ex)
            {
                _dispatcher.Invoke(() =>
                {
                    DeviceTextBlock.Text = "Error: " + ex.Message;
                });
            }
        }


        private void MainWindow_OnClosing(object sender, CancelEventArgs e)
        {
            _insertWatcher.Stop();
            _insertWatcher.Dispose();
            _deleteWatcher.Stop();
            _deleteWatcher.Dispose();
        }
    }
}
