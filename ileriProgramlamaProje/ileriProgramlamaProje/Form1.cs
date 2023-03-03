using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Media;
using Ex = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace ileriProgramlamaProje
{
    public partial class Form1 : Form
    {
        List<string> tekerlemeler = new List<string>();
        List<string> kelimeler = new List<string>();
        List<string> fıkralar = new List<string>();
        List<string> nefes = new List<string>();
        List<string> yanak = new List<string>();
        List<string> dil = new List<string>();
        List<string> dudak = new List<string>();
        List<string> cene = new List<string>();
        List<string> ses = new List<string>();

        List<string> liste;


        int sayac = 0, sure=0, secenek=0;
        int var = 0, index = 0;
        string suremetni="00:00:00";

        public Form1()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (sayac % 4 == 0)
            {
                pictureBox1.Image = ileriProgramlamaProje.Properties.Resources.resim2;

                pictureBox2.Image = ileriProgramlamaProje.Properties.Resources.noktam;
                pictureBox5.Image = ileriProgramlamaProje.Properties.Resources.noktab;

                label6.Text = "TEKERLEMELER";
                label7.Text = "Tekerleme söyleme dile işlevsellik kazandırır. Tekerleme söyle kısmından sizler için seçmiş olduğumuz tekerlemeleri okuyabilirsiniz.";
            }
            else if (sayac % 4 == 1)
            {
                pictureBox1.Image = ileriProgramlamaProje.Properties.Resources.resim1;

                pictureBox3.Image = ileriProgramlamaProje.Properties.Resources.noktam;
                pictureBox2.Image = ileriProgramlamaProje.Properties.Resources.noktab;
                label6.Text = "KELİMELER";
                label7.Text = "Kelime söyleme kelimelerin doğru telaffuz edebilmeleri için yarar sağlayacaktır. Kelime söyle kısmandan kelimeyle ilgili görsel ve seslendirmelere ulaşabilirsiniz.";
            }
            else if (sayac % 4 == 2)
            {
                pictureBox1.Image = ileriProgramlamaProje.Properties.Resources.resim3;

                pictureBox4.Image = ileriProgramlamaProje.Properties.Resources.noktam;
                pictureBox3.Image = ileriProgramlamaProje.Properties.Resources.noktab;

                label6.Text = "EGZERSİZLER";
                label7.Text = "Egzersiz yapma konuşma bozukluğunda iyileştirme sağlamaktadır. Egzersiz yapma kısmında sizler için hazırlayıp kategorize ettiğimiz egzersizleri yapabilirsiniz.";
            }
            else if (sayac % 4 == 3)
            {
                pictureBox1.Image = ileriProgramlamaProje.Properties.Resources.resim4;

                pictureBox5.Image = ileriProgramlamaProje.Properties.Resources.noktam;
                pictureBox4.Image = ileriProgramlamaProje.Properties.Resources.noktab;
                label6.Text = "FIKRALAR";
                label7.Text = "Fıkra okuma sizlere okuma alışkanlığı kazandırıp okumanızı hızlandıracaktır. Fıkra oku kısmından sizler için seçmiş olduğumuz fıkraları okuyabilirsiniz.";
            }

            sayac++;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;

            string temppath = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllBytes(temppath, ileriProgramlamaProje.Properties.Resources.tümveriler);

            Ex.Application uygulama = new Ex.Application();
            Ex.Workbook kitap = uygulama.Workbooks.Open(temppath);
            Ex.Worksheet sayfa  = kitap.Worksheets[1];
            Ex.Worksheet sayfa2 = kitap.Worksheets[2];
            Ex.Worksheet sayfa3 = kitap.Worksheets[3];
            Ex.Worksheet sayfa4 = kitap.Worksheets[4];
            Ex.Worksheet sayfa5 = kitap.Worksheets[5];
            Ex.Worksheet sayfa6 = kitap.Worksheets[6];
            Ex.Worksheet sayfa7 = kitap.Worksheets[7];
            Ex.Worksheet sayfa8 = kitap.Worksheets[8];
            Ex.Worksheet sayfa9 = kitap.Worksheets[9];

            int son = sayfa.UsedRange.Rows.Count;

            for (int i = 1; i <= son; i++)
            {
                tekerlemeler.Add(sayfa.Cells[i, 1].Value2);
            }

            int son2 = sayfa2.UsedRange.Rows.Count;

            for (int i = 1; i <= son2; i++)
            {
                kelimeler.Add(sayfa2.Cells[i, 1].Value2);
            }

            int son3 = sayfa3.UsedRange.Rows.Count;

            for (int i = 1; i <= son3; i++)
            {
                fıkralar.Add(sayfa3.Cells[i, 1].Value2);
            }


            int son4 = sayfa4.UsedRange.Rows.Count;

            for (int i = 1; i <= son4; i++)
            {
                cene.Add(sayfa4.Cells[i, 1].Value2);
            }

            int son5 = sayfa5.UsedRange.Rows.Count;

            for (int i = 1; i <= son5; i++)
            {
                yanak.Add(sayfa5.Cells[i, 1].Value2);
            }

            int son6 = sayfa6.UsedRange.Rows.Count;

            for (int i = 1; i <= son6; i++)
            {
                dudak.Add(sayfa6.Cells[i, 1].Value2);
            }

            int son7 = sayfa7.UsedRange.Rows.Count;

            for (int i = 1; i <= son7; i++)
            {
                dil.Add(sayfa7.Cells[i, 1].Value2);
            }

            int son8 = sayfa8.UsedRange.Rows.Count;

            for (int i = 1; i <= son8; i++)
            {
                nefes.Add(sayfa8.Cells[i, 1].Value2);
            }

            int son9 = sayfa9.UsedRange.Rows.Count;

            for (int i = 1; i <= son9; i++)
            {
                ses.Add(sayfa9.Cells[i, 1].Value2);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(sayfa);
            Marshal.ReleaseComObject(sayfa2);
            Marshal.ReleaseComObject(sayfa3);
            Marshal.ReleaseComObject(sayfa4);
            Marshal.ReleaseComObject(sayfa5);
            Marshal.ReleaseComObject(sayfa6);
            Marshal.ReleaseComObject(sayfa7);
            Marshal.ReleaseComObject(sayfa8);
            Marshal.ReleaseComObject(sayfa9);

            kitap.Close();
            Marshal.ReleaseComObject(kitap);

            uygulama.Quit();
            Marshal.ReleaseComObject(uygulama);

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = ileriProgramlamaProje.Properties.Resources.resim2;
            pictureBox2.Image = ileriProgramlamaProje.Properties.Resources.noktam;
            pictureBox3.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox4.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox5.Image = ileriProgramlamaProje.Properties.Resources.noktab;

            label6.Text = "TEKERLEMELER";
            label7.Text = "Tekerleme söyleme dile işlevsellik kazandırır. Tekerleme söyle kısmından sizler için seçmiş olduğumuz tekerlemeleri okuyabilirsiniz.";

            sayac = 0;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = ileriProgramlamaProje.Properties.Resources.resim1;
            pictureBox2.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox3.Image = ileriProgramlamaProje.Properties.Resources.noktam;
            pictureBox4.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox5.Image = ileriProgramlamaProje.Properties.Resources.noktab;

            label6.Text = "KELİMELER";
            label7.Text = "Kelime söyleme kelimelerin doğru telaffuz edebilmeleri için yarar sağlayacaktır. Kelime söyle kısmandan kelimeyle ilgili görsel ve seslendirmelere ulaşabilirsiniz.";

            sayac = 1;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = ileriProgramlamaProje.Properties.Resources.resim3;
            pictureBox2.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox3.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox4.Image = ileriProgramlamaProje.Properties.Resources.noktam;
            pictureBox5.Image = ileriProgramlamaProje.Properties.Resources.noktab;

            label6.Text = "EGZERSİZLER";
            label7.Text = "Egzersiz yapma konuşma bozukluğunda iyileştirme sağlamaktadır. Egzersiz yapma kısmında sizler için hazırlayıp kategorize ettiğimiz egzersizleri yapabilirsiniz.";

            sayac = 2;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = ileriProgramlamaProje.Properties.Resources.resim4;
            pictureBox2.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox3.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox4.Image = ileriProgramlamaProje.Properties.Resources.noktab;
            pictureBox5.Image = ileriProgramlamaProje.Properties.Resources.noktam;

            label6.Text = "FIKRALAR";
            label7.Text = "Fıkra okuma sizlere okuma alışkanlığı kazandırıp okumanızı hızlandıracaktır. Fıkra oku kısmından sizler için seçmiş olduğumuz fıkraları okuyabilirsiniz.";

            sayac = 3;
        }

        private void label2_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel6.Visible = false;
            panel4.Visible = true;
            tableLayoutPanel7.Visible = true;
            tableLayoutPanel12.Visible = false;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
            secenek = 1;
            label12.Text = $"Harf Seçiniz!";
            comboBox1.SelectedItem = null;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            panel4.Visible = false;
            panel6.Visible = false;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.anasayfa_ap;
        }

        private void label_MouseEnter(object sender, EventArgs e)
        {
            Label l = (Label) sender;
            l.BackColor = Color.WhiteSmoke;
        }

        private void label_MouseLeave(object sender, EventArgs e)
        {
            Label l = (Label)sender;
            l.BackColor = Color.Transparent;
        }

        private void label3_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel6.Visible = false;
            panel4.Visible = true;
            tableLayoutPanel7.Visible = true;
            pictureBox11.Image = null;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
            secenek = 2;
            label12.Text = $"Harf Seçiniz!";
            comboBox1.SelectedItem = null;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            suremetni="";

            if (sure / 3600 > 9)
            {
                suremetni += $"{sure / 3600}:";
            }
            else
            {
                suremetni += $"0{sure / 3600}:";
            }
            
            if((sure / 60) % 60 > 9)
            {
                suremetni += $"{(sure / 60) % 60}:";
            }
            else
            {
                suremetni += $"0{(sure / 60) % 60}:";
            }

            if (sure % 60 > 9)
            {
                suremetni += $"{sure % 60}";
            }
            else
            {
                suremetni += $"0{sure % 60}";
            }

            button1.Text = $"{suremetni}";

            sure++;
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            timer2.Enabled = true;
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            timer2.Enabled = false;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                var = 0;

                if (secenek == 1)
                {
                    liste = tekerlemeler;
                }
                else if (secenek == 2)
                {
                    liste = kelimeler;
                    tableLayoutPanel12.Visible = true;
                }
                else if (secenek == 3)
                {

                }

                for (int i = 0; i < liste.Count; i++)
                {
                    if (liste[i].Substring(0, 1) == Convert.ToString(comboBox1.SelectedItem) || comboBox1.SelectedItem == "Karışık")
                    {
                        var = 1;
                        index = i;
                        break;
                    }
                }

                if (var == 1)
                {

                    if (secenek == 2)
                    {
                        pictureBox11.Image = (Image)ileriProgramlamaProje.Properties.Resources.ResourceManager.GetObject(kelimeler[index]);
                    }

                    label12.Text = $"{liste[index]}";
                }
                else if (var == 0)
                {
                    if (secenek == 1)
                    {
                        label12.Text = $"{comboBox1.SelectedItem} ile başlayan tekerleme bulunamamıştır!";
                    }
                    else if (secenek == 2)
                    {
                        label12.Text = $"{comboBox1.SelectedItem} ile başlayan kelime bulunamamıştır!";
                    }
                    else if (secenek == 3)
                    {
                        label12.Text = $"{comboBox1.SelectedItem} ile başlayan egzersiz bulunamamıştır!";
                    }
                    else if (secenek == 4)
                    {
                        label12.Text = $"{comboBox1.SelectedItem} ile başlayan fıkra bulunamamıştır!";
                    }
                }

                label13.Text = label12.Text + "  🔊";
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel6.Visible = false;
            panel4.Visible = true;
            tableLayoutPanel7.Visible = false;
            tableLayoutPanel12.Visible = false;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
            secenek = 4;
            label12.Text = $"{fıkralar[0]}";
            index = 0;
            var = 1;
            liste = fıkralar;
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {
            Stream dosya = (Stream) ileriProgramlamaProje.Properties.Resources.ResourceManager.GetObject(kelimeler[index]+"1");
            SoundPlayer ses = new SoundPlayer(dosya);
            ses.Play();
        }

        private void pictureBox12_Click(object sender, EventArgs e)
        {
            SoundPlayer ses = new SoundPlayer(ileriProgramlamaProje.Properties.Resources.dogru);

            if (secenek == 4)
            {
                ses = new SoundPlayer(ileriProgramlamaProje.Properties.Resources.dogru2);
            }

            ses.Play();
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            SoundPlayer ses = new SoundPlayer(ileriProgramlamaProje.Properties.Resources.yanlis);

            if (secenek == 3)
            {
                ses = new SoundPlayer(ileriProgramlamaProje.Properties.Resources.yanlis2);
            }
            else if (secenek == 4)
            {
                ses = new SoundPlayer(ileriProgramlamaProje.Properties.Resources.yanlis3);
            }

            ses.Play();
        }


        private void button_MouseEnter(object sender, EventArgs e)
        {
            Button buton = (Button) sender;
            buton.BackgroundImage = ileriProgramlamaProje.Properties.Resources.butonap_3;
            buton.ForeColor = Color.Black;
        }

        private void button_MouseLeave(object sender, EventArgs e)
        {
            Button buton = (Button)sender;
            buton.BackgroundImage = ileriProgramlamaProje.Properties.Resources.butonap_1;
            buton.ForeColor = Color.White;
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {

        }

        private void button30_Click(object sender, EventArgs e)
        {
            liste = nefes;
            panel6.Visible = false;
            panel4.Visible = true;
            label12.Text = $"{liste[0]}";
            index = 0;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
        }

        private void button31_Click(object sender, EventArgs e)
        {
            liste = dil;
            panel6.Visible = false;
            panel4.Visible = true;
            label12.Text = $"{liste[0]}";
            index = 0;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            liste = yanak;
            panel6.Visible = false;
            panel4.Visible = true;
            label12.Text = $"{liste[0]}";
            index = 0;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
        }

        private void button37_Click(object sender, EventArgs e)
        {
            liste = cene;
            panel6.Visible = false;
            panel4.Visible = true;
            label12.Text = $"{liste[0]}";
            index = 0;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
        }

        private void button38_Click(object sender, EventArgs e)
        {
            liste = dudak;
            panel6.Visible = false;
            panel4.Visible = true;
            label12.Text = $"{liste[0]}";
            index = 0;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
        }

        private void button39_Click(object sender, EventArgs e)
        {
            liste = ses;
            panel6.Visible = false;
            panel4.Visible = true;
            label12.Text = $"{liste[0]}";
            index = 0;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.ekran2;
        }

        private void label15_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedItem = "Karışık";
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult cikis = new DialogResult();
            cikis = MessageBox.Show($"Bugün {suremetni} kadar çalıştınız. Uygulamadan çıkmak istiyor musunuz?", "Uyarı", MessageBoxButtons.YesNo);
            if (cikis == DialogResult.Yes)
            {
                Application.Exit();
            }
            if (cikis == DialogResult.No)
            {
                MessageBox.Show("Program kapatılmadı.");
                e.Cancel = true;

            }

        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null || secenek == 4 || secenek == 3)
            {
                int girdi = 0;

                string harf = Convert.ToString(comboBox1.SelectedItem);

                if (var == 1)
                {
                    for (int i = index - 1; i >= 0; i--)
                    {
                        if (liste[i].Substring(0, 1) == harf || secenek == 4 || secenek == 3 || harf == "Karışık")
                        {
                            girdi = 1;
                            index = i;
                            if (secenek == 2)
                            {
                                pictureBox11.Image = (Image)ileriProgramlamaProje.Properties.Resources.ResourceManager.GetObject(kelimeler[index]);
                            }
                            label12.Text = $"{liste[index]}";
                            break;
                        }
                    }

                    if (girdi == 0)
                    {
                        if (secenek == 1)
                        {
                            MessageBox.Show($"İlk tekerleme bu. Daha fazla geri gidemezsin! :)");
                        }
                        else if (secenek == 2)
                        {
                            MessageBox.Show($"İlk kelime bu. Daha fazla geri gidemezsin! :)");
                        }
                        else if (secenek == 3)
                        {
                            MessageBox.Show($"İlk egzersiz bu. Daha fazla geri gidemezsin! :)");
                        }
                        else if (secenek == 4)
                        {
                            MessageBox.Show($"İlk fıkra bu. Daha fazla geri gidemezsin! :)");
                        }

                    }
                }
                else if (var == 0)
                {
                    if (secenek == 1)
                    {
                        MessageBox.Show($"Maalesef {harf} harfiyle başlayan tekerleme bulunmamakta. Diğer harfleri deneyebilirsin! :)");
                    }
                    else if (secenek == 2)
                    {
                        MessageBox.Show($"Maalesef {harf} harfiyle başlayan kelime bulunmamakta. Diğer harfleri deneyebilirsin! :)");
                    }

                }

                label13.Text = label12.Text + "  🔊";

            }
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null || secenek == 4 || secenek == 3)
            {
                int girdi = 0;

                string harf = Convert.ToString(comboBox1.SelectedItem);

                if (var == 1)
                {
                    for (int i = index + 1; i < liste.Count; i++)
                    {
                        if (liste[i].Substring(0, 1) == harf || secenek == 4 || secenek == 3 || harf == "Karışık")
                        {
                            girdi = 1;
                            index = i;
                            if (secenek == 2)
                            {
                                pictureBox11.Image = (Image)ileriProgramlamaProje.Properties.Resources.ResourceManager.GetObject(kelimeler[index]);
                            }
                            label12.Text = $"{liste[index]}";
                            break;
                        }
                    }

                    if (girdi == 0)
                    {
                        if (secenek == 1)
                        {
                            if (harf == "Karışık")
                            {
                                MessageBox.Show($"Maalesef yeni tekerleme kalmadı. Başa dönebilirsin! :)");
                            }
                            else
                            {
                                MessageBox.Show($"Maalesef {harf} harfiyle başlayan yeni tekerleme kalmadı. Diğer harfleri deneyebilirsin! :)");
                            }
                        }
                        else if (secenek == 2)
                        {
                            if (harf == "Karışık")
                            {
                                MessageBox.Show($"Maalesef yeni kelime kalmadı. Başa dönebilirsin! :)");
                            }
                            else
                            {
                                MessageBox.Show($"Maalesef {harf} harfiyle başlayan yeni kelime kalmadı. Diğer harfleri deneyebilirsin! :)");
                            }
                        }
                        else if (secenek == 3)
                        {
                            MessageBox.Show($"Maalesef yeni egzersiz kalmadı. Diğer egzersiz türlerini deneyebilirsin! :)");
                        }
                        else if (secenek == 4)
                        {
                            MessageBox.Show($"Maalesef yeni tekerleme kalmadı. Başa dönüp tekrar okuyabilirsin! :)");
                        }
                    }
                }
                else if (var == 0)
                {
                    if (secenek == 1)
                    {
                        MessageBox.Show($"Maalesef {harf} harfiyle başlayan tekerleme bulunmamakta. Diğer harfleri deneyebilirsin! :)");
                    }
                    else if (secenek == 2)
                    {
                        MessageBox.Show($"Maalesef {harf} harfiyle başlayan kelime bulunmamakta. Diğer harfleri deneyebilirsin! :)");
                    }
                }

                label13.Text = label12.Text + "  🔊";
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel4.Visible = false;
            panel6.Visible = true;
            tableLayoutPanel7.Visible = false;
            tableLayoutPanel12.Visible = false;
            panel1.BackgroundImage = ileriProgramlamaProje.Properties.Resources.anasayfa_ap;
            secenek = 3;
            index = 0;
            var = 1;
        }
    }
}
