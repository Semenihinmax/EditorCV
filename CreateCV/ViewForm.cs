using System.Drawing;
using System.Windows.Forms;

namespace CreateCV
{
    public partial class ViewForm : Form
    {
        public ViewForm()
        {
            InitializeComponent();
        }

        public void getRichBox(IDataObject dataObject)
        {
            richTextBox.Rtf = dataObject.GetData(DataFormats.Rtf).ToString();
        }

        public void getPhoto(string photoAdress)
        {
            Image img = Image.FromFile(photoAdress);
            getImage(img);
        }

        public void getImage(Image img)
        {
            Clipboard.Clear();
            Clipboard.SetImage(img);
            richTextBox.Paste();
            Clipboard.Clear();
        }
    }
}
