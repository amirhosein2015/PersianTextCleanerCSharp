using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace PersianTextCleaner
{
    public partial class Form1 : Form
    {

        string openedText = "";


        public Form1()
        {
            InitializeComponent();


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "All Supported Files|*.txt;*.doc;*.docx|Text Files|*.txt|Word Documents|*.doc;*.docx"
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor; //  موس حالت انتظار
                string filePath = dlg.FileName;
                string ext = Path.GetExtension(filePath).ToLower();

                try
                {
                    if (ext == ".txt")
                    {
                        openedText = File.ReadAllText(filePath, Encoding.UTF8);
                    }
                    else if (ext == ".doc" || ext == ".docx")
                    {
                        openedText = ReadWordFile(filePath);
                    }

                    textBoxOriginal.Text = openedText;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                }

                finally
                {
                    Cursor.Current = Cursors.Default; //  موس به حالت عادی برمی‌گردد
                }

            }
        }

        private string ReadWordFile(string filePath)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;
            string text = "";

            try
            {
                object path = filePath;
                object readOnly = true;
                object missing = System.Reflection.Missing.Value;

                doc = wordApp.Documents.Open(ref path, ReadOnly: ref readOnly, Visible: false);
                text = doc.Content.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading file Word: " + ex.Message);
            }
            finally
            {
                if (doc != null) doc.Close(false);
                wordApp.Quit();
            }

            return text;
        }








        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBoxCleaned.Text))
            {
                Clipboard.SetText(textBoxCleaned.Text);
                MessageBox.Show("text copied!");
            }
            else
            {
                MessageBox.Show("There is no content to copy.");
            }
        }



        private void buttonRun_Click(object sender, EventArgs e)
        {
            string originalText = textBoxOriginal.Text;

            if (!string.IsNullOrWhiteSpace(originalText))
            {
                string rewritten = RewriteStructure(originalText);
                textBoxCleaned.Text = rewritten;
            }
            else
            {
                MessageBox.Show("Please enter some text or open the file first.");
            }

        }

















        private void buttonSaveAs_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "Text File (*.txt)|*.txt|Word File (*.docx)|*.docx",
                FileName = "cleaned_text"
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string filePath = dlg.FileName;
                string ext = Path.GetExtension(filePath).ToLower();

                try
                {
                    if (ext == ".txt")
                    {
                        File.WriteAllText(filePath, textBoxCleaned.Text, Encoding.UTF8);
                    }
                    else if (ext == ".docx")
                    {
                        SaveToWord(filePath, textBoxCleaned.Text);
                    }

                    MessageBox.Show("The file was saved successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error saving file: " + ex.Message);
                }
            }
        }

        private void SaveToWord(string filePath, string content)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Add();

            try
            {
                doc.Content.Text = content;
                doc.SaveAs2(filePath);
            }
            finally
            {
                doc.Close();
                wordApp.Quit();
            }
        }







        private string RewriteStructure(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "";

            input = RemoveEmojis(input);

            string normalized = NormalizeText(input);
            string fixedLinguistic = ApplyLinguisticFixes(normalized);
            fixedLinguistic = FixPersianCompoundWords(fixedLinguistic);

            // باید اینجا بیاد
            fixedLinguistic = FixHehaPlural(fixedLinguistic);

            string final = FixPunctuationAndStructure(fixedLinguistic);
            return final.Trim();
        }




        private string NormalizeText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            text = RemoveArabicDiacritics(text);
            // ۱. حذف کاراکترهای کنترل نامرئی
            string[] invisibles = {
        "\u200e", "\u200f", "\u202a", "\u202b", "\u202c",
        "\u2066", "\u2067", "\u2069", "\u200c", "\u200d", "\ufeff"
    };
            foreach (var ch in invisibles)
                text = text.Replace(ch, "");

            // ۲. جایگزینی حروف عربی با معادل فارسی
            text = text.Replace('ي', 'ی')
                       .Replace('ك', 'ک')
                       .Replace('ة', 'ه')
                       .Replace('ؤ', 'و');

            // ۳. حذف کشیدگی‌ها (مثلاً: شکــــر → شکر)
            text = Regex.Replace(text, "([اآبپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی])\\1{2,}", "$1");

            // ۴. حذف فاصله‌های اضافی و یکنواخت‌سازی فاصله‌ها
            text = Regex.Replace(text, @"\s{2,}", " ");
            text = Regex.Replace(text, @"(\r?\n){3,}", "\n\n");

            // ۵. یکنواخت‌سازی علائم نگارشی فارسی و انگلیسی
            text = text.Replace(",", "،")
                       .Replace(";", "؛")
                       .Replace("?", "؟");

            return text.Trim();
        }



        private string ApplyLinguisticFixes(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            //// نیم‌فاصله
            //text = Regex.Replace(text, @"\b(می|نمی) ", "$1‌"); // می‌رود / نمی‌خواهد




            // ترکیب‌های رایج به نیم‌فاصله
            text = Regex.Replace(text, @"\b(توصیه|پیشنهاد|احساس|نظر|دلیل|هدف|درخواست|تجربه) (من|ما|تو|شما|او|آنها)\b", "$1‌$2");

            // معادل‌سازی نوشتاری برای واژه‌های عامیانه رایج
            text = text.Replace("اگه", "اگر")
                       .Replace("بخاطر", "به‌خاطر")
                       .Replace("نمیخوام", "نمی‌خوام")
                       .Replace("میخواد", "می‌خواد")
                       .Replace("نمیتونه", "نمی‌تونه")
                       .Replace("میتونه", "می‌تونه") //بعد-قبل
                       .Replace("خدارو", "خدا رو")
                       .Replace("خداروشکر", "خدا را شکر");

            // اصلاح شکل گفتاری پرکاربرد
            text = Regex.Replace(text, @"\bیه\b", "یک");
            text = Regex.Replace(text, @"\bچراکه\b", "چرا که");
            text = Regex.Replace(text, @"\bبعدشم\b", "بعدش هم");

      




            // --- بعد از تمام اصلاحات قبلی در ApplyLinguisticFixes اینو اضافه کن ---
            var wordFixes = new Dictionary<string, string>
{

    { "میتوانم", "می‌توانم" }, { "نمیتوانم", "نمی‌توانم" },
        { "میدانم", "می‌دانم" }, { "نمیخواهم", "نمی‌خواهم" },
        { "میخواهم", "می‌خواهم" }, { "میروم", "می‌روم" },
        { "میایم", "می‌آیم" }, { "میآیم", "می‌آیم" },
        { "نمیایم", "نمی‌آیم" }, { "میکنم", "می‌کنم" },
        { "نمیکنم", "نمی‌کنم" }, { "میگیرم", "می‌گیرم" },
        { "میبینم", "می‌بینم" }, { "نمیبینم", "نمی‌بینم" },
        { "میگویم", "می‌گویم" }, { "میشنوم", "می‌شنوم" },
        { "نمیشنوم", "نمی‌شنوم" }, { "میدهم", "می‌دهم" },
        { "نمیروم", "نمی‌روم" }, { "نمیخواستم", "نمی‌خواستم" },
        { "میرفت", "می‌رفت" }, { "نمیرفت", "نمی‌رفت" },
        { "میآمد", "می‌آمد" }, { "میرفتیم", "می‌رفتیم" },
        { "میرفتند", "می‌رفتند" }, { "میآورد", "می‌آورد" },
        { "میکرد", "می‌کرد" }, { "نمیکرد", "نمی‌کرد" },
        { "نمیخواستی", "نمی‌خواستی" }, { "میخواهد", "می‌خواهد" },
        { "نمیخواهد", "نمی‌خواهد" }, { "میپرسد", "می‌پرسد" },
        { "میگذارم", "می‌گذارم" }, { "نمیگذارم", "نمی‌گذارم" },
        { "نمیگذارند", "نمی‌گذارند" }, { "میگذارند", "می‌گذارند" },

        // افعال محاوره‌ای پرکاربرد
        { "میدونم", "می‌دونم" }, { "نمیدونم", "نمی‌دونم" },
        { "میدونی", "می‌دونی" }, { "نمیدونی", "نمی‌دونی" },
        { "میدونه", "می‌دونه" }, { "نمیدونه", "نمی‌دونه" },
        { "میدونیم", "می‌دونیم" }, { "نمیدونیم", "نمی‌دونیم" },
        { "میدونید", "می‌دونید" }, { "نمیدونید", "نمی‌دونید" },
        { "میدونن", "می‌دونن" }, { "نمیدونن", "نمی‌دونن" },

        { "میخندم", "می‌خندم" }, { "نمیخندم", "نمی‌خندم" },
        { "میخندی", "می‌خندی" }, { "نمیخندی", "نمی‌خندی" },
        { "میخنده", "می‌خنده" }, { "نمیخنده", "نمی‌خنده" },

        { "میخواستم", "می‌خواستم" }, { "میخواستی", "می‌خواستی" },
        { "میخواست", "می‌خواست" }, { "میخواستیم", "می‌خواستیم" },
        { "میخواستید", "می‌خواستید" }, { "میخواستند", "می‌خواستند" },

        { "میتونم", "می‌تونم" }, { "نمیتونم", "نمی‌تونم" },
        { "میتونی", "می‌تونی" }, { "نمیتونی", "نمی‌تونی" },
        { "میتونه", "می‌تونه" }, { "نمیتونه", "نمی‌تونه" },
        { "میتونیم", "می‌تونیم" }, { "نمیتونیم", "نمی‌تونیم" },
        { "میتونید", "می‌تونید" }, { "نمیتونید", "نمی‌تونید" },
        { "میتونن", "می‌تونن" }, { "نمیتونن", "نمی‌تونن" },

        { "میگم", "می‌گم" }, { "نمیگم", "نمی‌گم" },
        { "میگی", "می‌گی" }, { "نمیگی", "نمی‌گی" },
        { "میگه", "می‌گه" }, { "نمیگه", "نمی‌گه" },
        { "میگیم", "می‌گیم" }, { "نمیگیم", "نمی‌گیم" },
        { "میگید", "می‌گید" }, { "نمیگید", "نمی‌گید" },
        { "میگن", "می‌گن" }, { "نمیگن", "نمی‌گن" },

        { "میگویی", "می‌گویی" }, { "نمیگویی", "نمی‌گویی" },
        { "میگوید", "می‌گوید" }, { "نمیگوید", "نمی‌گوید" },
        { "میگوییم", "می‌گوییم" }, { "نمیگوییم", "نمی‌گوییم" },
        { "میگویید", "می‌گویید" }, { "نمیگویید", "نمی‌گویید" },
        { "میگویند", "می‌گویند" }, { "نمیگویند", "نمی‌گویند" },



{ "میخورم", "می‌خورم" }, { "نمیخورم", "نمی‌خورم" },
{ "میخوری", "می‌خوری" }, { "نمیخوری", "نمی‌خوری" },
{ "میخوره", "می‌خوره" }, { "نمیخوره", "نمی‌خوره" },

{ "میریم", "می‌ریم" }, { "نمیریم", "نمی‌ریم" },
{ "میری", "می‌ری" }, { "نمیری", "نمی‌ری" },
{ "میره", "می‌ره" }, { "نمیره", "نمی‌ره" },
{ "میان", "می‌آن" }, { "نمیان", "نمی‌آن" },

{ "میذارم", "می‌ذارم" }, { "نمیذارم", "نمی‌ذارم" },
{ "میذاره", "می‌ذاره" }, { "نمیذاره", "نمی‌ذاره" },

{ "میشنوی", "می‌شنوی" }, { "نمیشنوی", "نمی‌شنوی" },
{ "میشنوه", "می‌شنوه" }, { "نمیشنوه", "نمی‌شنوه" },

{ "میبرم", "می‌برم" }, { "نمیبرم", "نمی‌برم" },
{ "میبری", "می‌بری" }, { "نمیبری", "نمی‌بری" },
{ "میبره", "می‌بره" }, { "نمیبره", "نمی‌بره" },

{ "میذارید", "می‌ذارید" }, { "نمیذارید", "نمی‌ذارید" },
{ "میذاریم", "می‌ذاریم" }, { "نمیذاریم", "نمی‌ذاریم" },
{ "میذارن", "می‌ذارن" }, { "نمیذارن", "نمی‌ذارن" },

{ "میبینیش", "می‌بینیش" }, { "نمیبینیش", "نمی‌بینیش" },
{ "میخوامش", "می‌خوامش" }, { "نمیخوامش", "نمی‌خوامش" },
{ "میگیرمش", "می‌گیرمش" }, { "نمیگیرمش", "نمی‌گیرمش" },



// فعل "زدن"
{ "میزنم", "می‌زنم" }, { "نمیزنم", "نمی‌زنم" },
{ "میزنه", "می‌زنه" }, { "نمیزنه", "نمی‌زنه" },
{ "میزنی", "می‌زنی" }, { "نمیزنی", "نمی‌زنی" },


// فعل "فهمیدن"
{ "میفهمم", "می‌فهمم" }, { "نمیفهمم", "نمی‌فهمم" },
{ "میفهمی", "می‌فهمی" }, { "نمیفهمی", "نمی‌فهمی" },
{ "میفهمه", "می‌فهمه" }, { "نمیفهمه", "نمی‌فهمه" }


};

            foreach (var pair in wordFixes)
            {
                text = Regex.Replace(text, $@"\b{pair.Key}\b", pair.Value);
            }


            return text.Trim();
        }




        private string FixPunctuationAndStructure(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            // حذف کشیدگی‌های بیش‌ازحد حروف (مثلاً ســـــلام → سلام)
            text = Regex.Replace(text, "([اآبپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی])\\1{2,}", "$1");

            // اصلاح فاصله قبل و بعد از علائم نگارشی
            text = Regex.Replace(text, @"\s*([!.,؛:؟…])\s*", "$1 ");  // فاصله قبل حذف، بعد اضافه
            text = Regex.Replace(text, @"\s{2,}", " ");               // حذف فاصله‌های تکراری

            // یکدست‌سازی علائم نگارشی
            text = text.Replace("...", "…");
            text = text.Replace("..", ".");
            text = text.Replace("!!", "!");
            text = text.Replace("؟؟", "؟");
            text = Regex.Replace(text, @"([!.؟]){2,}", "$1"); // اصلاح علائم پشت سر هم

            // پایان هر جمله با نقطه یا علامت مناسب
            text = Regex.Replace(text, @"([^\n])\n", "$1.\n"); // اگر خط بدون نقطه تموم شده
            text = Regex.Replace(text, @"(?<![.!؟])(\s|$)", "  "); // پایان جمله بدون نقطه




            // حذف فاصله قبل از نقطه پایان
            text = Regex.Replace(text, @" \.", ".");

            // اصلاح فاصله‌های قبل و بعد از پرانتز
            text = Regex.Replace(text, @"\(\s+", "(");
            text = Regex.Replace(text, @"\s+\)", ")");

            // حذف خطوط خالی بیشتر از ۲ تا
            text = Regex.Replace(text, @"(\r?\n){3,}", "\n\n");



            // حذف فاصله‌های اضافی اول و آخر متن
            return text.Trim();

        }





        private string FixPersianCompoundWords(string text)
        {
            //string[] prefixes = { "می", "نمی" };

            //foreach (var prefix in prefixes.OrderByDescending(p => p.Length))
            //{
            //    var pattern = $@"(?<!\S){prefix}(?=[آ-یءئؤا-ي]{{2,}})";
            //    text = Regex.Replace(text, pattern, $"{prefix}\u200c", RegexOptions.IgnoreCase);
            //}

            return text;
        }



        public static string RemoveEmojis(string input)
        {
            return Regex.Replace(input, @"[\uD800-\uDBFF][\uDC00-\uDFFF]", "");
        }




        public string RemoveArabicDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            // لیست حرکات مزاحم فتحه و کسره بدون همزه ء)
            string diacritics = "\u064B\u064C\u064D\u064E\u064F\u0650\u0651\u0652";

            foreach (char c in diacritics)
            {
                text = text.Replace(c.ToString(), "");
            }

            return text;
        }







        private string FixHehaPlural(string input)
        {
            // نیم‌فاصله یونیکد
            string zwnj = "\u200c";

            // اصلاح مواردی مثل: خانهها -> خانه‌ها
            return Regex.Replace(input, @"(?<=[\u0647])ها\b", zwnj + "ها");
        }



        private void textBoxCleaned_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            textBoxOriginal.Clear();
            textBoxCleaned.Clear();
            openedText = string.Empty;
            MessageBox.Show("All fields have been cleared.");
        }
    }
}