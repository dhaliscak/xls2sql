using System.Text;
using System.Windows;

namespace xls2sql
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
#if NET5_0_OR_GREATER
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
            base.OnStartup(e);
        }
    }
}
