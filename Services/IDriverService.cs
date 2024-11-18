using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace FormFiller.Services
{
    public interface IDriverService
    {
        (ChromeDriver, WebDriverWait) GetDriver();
    }

    public class DriverService : IDriverService, IDisposable
    {
        private static DriverService _instance;
        private static readonly object _lock = new object();

        private ChromeDriver _driver;
        private WebDriverWait _wait;

        private bool _disposed = false;

        private DriverService()
        {
            InitializeDriver();
        }

        // Singleton instance property
        public static DriverService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new DriverService();
                        }
                    }
                }
                return _instance;
            }
        }

        private void InitializeDriver()
        {
            try
            {
                var options = new ChromeOptions();
                _driver = new ChromeDriver(options);
                _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        public (ChromeDriver, WebDriverWait) GetDriver()
        {
            if (_driver == null || _disposed)
            {
                InitializeDriver();
            }
            return (_driver, _wait);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _driver?.Quit();
                    _driver?.Dispose();
                    _wait = null;
                }
                _disposed = true;
            }
        }
    }
}
