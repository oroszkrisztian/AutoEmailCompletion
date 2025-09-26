using System.Configuration;
using System.Data;
using System.Windows;
using System.Text;

namespace EmailCompleteApp
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
