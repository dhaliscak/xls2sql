using System.Text;
using System.Windows;

namespace xls2sql
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            base.OnStartup(e);
        }
    }
}
