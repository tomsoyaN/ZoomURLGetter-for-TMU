using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ZoomURLGetter
{
    /// <summary>
    /// SubWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class SubWindow : Window
    {
        public SubWindow(string txts)
        {
            InitializeComponent();
            Paragraph par = new Paragraph();
            Run r = new Run(txts);
            par.Inlines.Add(r);
            TB.Document.Blocks.Add(par);
        }
    }
}
