using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace CreateCV
{
    public partial class CVForm : Form
    {
        public CVForm()
        {
            InitializeComponent();
        }

        private string imageAdress;
        private Image image;

        private void createCVButton_Click(object sender, EventArgs e)
        {
            if (checkTextBoxs())
            {
                var helper = new WordHelper("PatternCV.docx");

                var items = new Dictionary<string, string>
                {
                {"<FIO>", fioTextBox.Text},
                {"<PHONE>", phoneTextBox.Text},
                {"<EMAIL>", emailTextBox.Text},
                {"<EDUCATION>", educationTextBox.Text},
                {"<WORK_EXPERIENCE>", experienceTextBox.Text},
                {"<SKILLS>", skillsTextBox.Text},
                {"<PERSONAL_ACHIEVEMENTS>", achievementsTextBox.Text},
                {"<HOBBY>", hobbyTextBox.Text}
                };

                helper.Process(items, imageAdress);
            }
        }

        Regex regFio = new Regex(@"^(?<FIO>.+)\S\r$", RegexOptions.Compiled | RegexOptions.Singleline);//don't work
        Regex regPhone = new Regex(@"Телефон: (?<PHONE>.+)\r", RegexOptions.Compiled | RegexOptions.Singleline);
        Regex regEmail = new Regex(@"Email: (?<EMAIL>.+)", RegexOptions.Compiled | RegexOptions.Singleline);

        private void loadCVButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Word 97-2003|*.doc|Word Document|*.docx";
            if (openDialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                var helper = new WordHelper(openDialog.FileName);
                List<string> textBlocks = helper.ReadText();
                image = helper.ReadImage();

                string fio = null;
                string phone = null;
                string email = null;
                string education = null;
                string experience = null;
                string skills = null;
                string achievements = null;
                string hobby = null;

                bool educationState = false;
                bool experienceState = false;
                bool skillsState = false;
                bool achievementsState = false;
                bool hobbyState = false;

                foreach (var block in textBlocks.Where(b => !string.IsNullOrEmpty(b)))
                {
                    Match match = null;
                    if (string.IsNullOrEmpty(fio))
                    {
                        match = regFio.Match(block);
                        if( match != null && match.Success)
                        {
                            fio = match.Groups["FIO"].Value;
                        }
                    }

                    if (string.IsNullOrEmpty(phone))
                    {
                        match = regPhone.Match(block);
                        if (match != null && match.Success)
                        {
                            phone = match.Groups["PHONE"].Value;
                        }
                    }

                    if (string.IsNullOrEmpty(email))
                    {
                        match = regEmail.Match(block);
                        if (match != null && match.Success)
                        {
                            email = match.Groups["EMAIL"].Value;
                        }
                    }

                    if (educationState)
                    {
                        education += block.ToString() != "Опыт работы\r" ? block.ToString() : "";
                    }

                    if (experienceState)
                    {
                        experience += block.ToString() != "Навыки\r" ? block.ToString() : "";
                    }

                    if (skillsState)
                    {
                        skills += block.ToString() != "Личные достижения\r" ? block.ToString() : "";
                    }

                    if (achievementsState)
                    {
                        achievements += block.ToString() != "Хобби\r" ? block.ToString() : "";
                    }

                    if (hobbyState)
                    {
                        hobby += block.ToString();
                    }

                    if (block.ToString() == "Опыт работы\r")
                    {
                        educationState = false;
                        experienceState = true;
                    }

                    if (block.ToString() == "Образование\r")
                    {
                        educationState = true;
                    }

                    if (block.ToString() == "Навыки\r")
                    {
                        experienceState = false;
                        skillsState = true;
                    }

                    if (block.ToString() == "Личные достижения\r")
                    {
                        skillsState = false;
                        achievementsState = true;
                    }

                    if (block.ToString() == "Хобби\r")
                    {
                        achievementsState = false;
                        hobbyState = true;
                    }
                }

                if (image != null)
                {
                    infoPhotoLabel.Text = openDialog.FileName;
                    seePhotoButton.Visible = true;
                }

                fioTextBox.Text = fio;
                phoneTextBox.Text = phone;
                emailTextBox.Text = email;
                educationTextBox.Text = education.Replace("\r","\r\n");
                experienceTextBox.Text = experience.Replace("\r", "\r\n");
                skillsTextBox.Text = skills.Replace("\r", "\r\n");
                achievementsTextBox.Text = achievements.Replace("\r", "\r\n");
                hobbyTextBox.Text = hobby.Replace("\r", "\r\n");
            }
            catch { }
        }

        private void phoneTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && (e.KeyChar != '+') && (e.KeyChar != 8))
            {
                e.Handled = true;
            }
        }

        private void fioTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Client can use Latin and Russian symbols
            if ((e.KeyChar <= 1039 || e.KeyChar >= 1104) && (e.KeyChar <= 96 || e.KeyChar >= 123) &&  (e.KeyChar <= 64 || e.KeyChar >= 91) && (e.KeyChar != 8))
            {
                e.Handled = true;
            }
        }

        private void viewCVButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                ValidateNames = true,
                Multiselect = false,
                Filter = "Word 97-2003|*.doc|Word Document|*.docx"
            })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    object readOnly = false;
                    object visible = true;
                    object save = false;
                    object fileName = openFileDialog.FileName;
                    object newTemplate = false;
                    object docType = 0;
                    object missing = Type.Missing;

                    Microsoft.Office.Interop.Word.Document document;
                    Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application() { Visible = false };

                    document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref visible, ref missing, ref missing, ref missing, ref missing);
                    document.ActiveWindow.Selection.WholeStory();
                    document.ActiveWindow.Selection.Copy();
                    
                    IDataObject dataObject = Clipboard.GetDataObject();

                    ViewForm viewForm = new ViewForm();
                    viewForm.getRichBox(dataObject);
                    viewForm.Show();
                    
                    application.Quit(ref missing, ref missing, ref missing);
                }
            }
        }

        private void loadPhotoButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файлы изображений|*.bmp;*.png;*.jpg";
            if (openDialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                imageAdress = openDialog.FileName;
                infoPhotoLabel.Text = imageAdress.Split(new char[] { '\\' }).Last();
                seePhotoButton.Visible = true;
            }
            catch (OutOfMemoryException ex)
            {
                MessageBox.Show("Ошибка чтения картинки");
                return;
            }
        }

        private bool checkTextBoxs()
        {
            if (fioTextBox.Text.Trim() == "")
            {
                MessageBox.Show("Введите ФИО");
                return false;
            }
            if (phoneTextBox.Text.Trim() == "" && emailTextBox.Text.Trim() == "")
            {
                MessageBox.Show("Введите почту или телефон");
                return false;
            }
            if (educationTextBox.Text.Trim() == "")
            {
                MessageBox.Show("Введите данные об образовании");
                return false;
            }
            if (skillsTextBox.Text.Trim() == "")
            {
                MessageBox.Show("Введите информацию о навыках");
                return false;
            }

            return true;
        }

        private void seePhotoButton_Click(object sender, EventArgs e)
        {
            ViewForm viewForm = new ViewForm();

            if(image == null)
            {
                viewForm.getPhoto(imageAdress);
            }
            else
            {
                viewForm.getImage(image);
            }

            viewForm.Show();
        }
    }
}
