using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.Drawing;
using System.IO;
using System.Net;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
// берем date-time штамп
using System.Globalization;
using System.Threading;
// для работы с Excel
using OfficeOpenXml;

namespace Расчёт_Окамура_Хата_Ли
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            // предварительные подстановки числовых данных
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле M
            this.textBox1.Text = "8";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле P(bs) (мВт)
            this.textBox2.Text = "100000";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле g
            this.textBox4.Text = "0";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле h1(м)
            this.textBox5.Text = "15";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле h1(м)
            this.textBox6.Text = "25";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле h1(м)
            this.textBox7.Text = "35";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле h1(м)
            this.textBox8.Text = "50";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле h1(м)
            this.textBox9.Text = "75";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле h2(м)
            this.textBox10.Text = "3";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Y
            this.textBox11.Text = "48";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле dh
            this.textBox12.Text = "60";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле f
            this.textBox13.Text = "400000000";
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле L
            this.textBox14.Text = "2";
            ///////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////
            ///
            /// Хата - R
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Хата - R10
            this.textBox46.Text = "10";
            this.textBox46.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Хата - R20
            this.textBox47.Text = "20";
            this.textBox47.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Хата - R30
            this.textBox48.Text = "30";
            this.textBox48.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Хата - R 40
            this.textBox49.Text = "40";
            this.textBox49.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Хата - R 50
            this.textBox50.Text = "50";
            this.textBox50.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Хата - R 60
            this.textBox51.Text = "60";
            this.textBox51.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////
            ///
            /// Ли - R
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Ли - R10
            this.textBox82.Text = "10";
            this.textBox82.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Ли - R20
            this.textBox83.Text = "20";
            this.textBox83.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Ли - R30
            this.textBox84.Text = "30";
            this.textBox84.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Ли - R 40
            this.textBox85.Text = "40";
            this.textBox85.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Ли - R 50
            this.textBox86.Text = "50";
            this.textBox86.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Ли - R 60
            this.textBox87.Text = "60";
            this.textBox87.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////
            ///
            /// Окамура - R
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Окамура - R10
            this.textBox118.Text = "10";
            this.textBox118.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Окамура - R20
            this.textBox119.Text = "20";
            this.textBox119.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Окамура - R30
            this.textBox120.Text = "30";
            this.textBox120.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Окамура - R 40
            this.textBox121.Text = "40";
            this.textBox121.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Окамура - R 50
            this.textBox122.Text = "50";
            this.textBox122.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Окамура - R 60
            this.textBox123.Text = "60";
            this.textBox123.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////

            // расчет P(bs) (дБм) при появлении данных в P(bs) (мВт)
            // готовим поля для работы - P(bs) (дБм)
            double val = 10 * Math.Log10(Convert.ToDouble(this.textBox2.Text) / 1);
            this.textBox3.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле - P(bs) (дБм)
            this.textBox3.BackColor = Color.FromArgb(192, 255, 192);

            //////////////////////////////////////////////////////

            // расчет Длина волны при появлении данных в f
            // готовим поля для работы - Длина волны
            val = 300000000 / Convert.ToDouble(this.textBox13.Text);
            this.textBox15.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле - Длина волны
            this.textBox15.BackColor = Color.FromArgb(192, 255, 192);

            ///
            //////////////////////////////////////////////////////
            /// Текстовые данные полей таблицы Am (табл. 3.1)

            // ставим текстовое значение индикатор на поле Am (табл. 3.1) - 15
            this.textBox124.Text = "15";
            this.textBox124.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Am (табл. 3.1) - 18
            this.textBox125.Text = "18";
            this.textBox125.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Am (табл. 3.1) - 21
            this.textBox126.Text = "21";
            this.textBox126.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Am (табл. 3.1) - 25
            this.textBox127.Text = "25";
            this.textBox127.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Am (табл. 3.1) - 28
            this.textBox128.Text = "28";
            this.textBox128.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            //////
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле Am (табл. 3.1) - 31
            this.textBox129.Text = "31";
            this.textBox129.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///
            //////////////////////////////////////////////////////
            /// Текстовые данные полей усреденных данных 3х моделей

            // ставим текстовое значение индикатор на поле усреденных данных 3х моделей - 5
            this.textBox136.Text = "5";
            this.textBox136.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле усреденных данных 3х моделей - 10
            this.textBox137.Text = "10";
            this.textBox137.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле усреденных данных 3х моделей - 15
            this.textBox138.Text = "15";
            this.textBox138.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле усреденных данных 3х моделей - 20
            this.textBox139.Text = "20";
            this.textBox139.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            ///
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле усреденных данных 3х моделей - 25
            this.textBox140.Text = "25";
            this.textBox140.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
            //////
            ///////////////////////////////////////////////////////////
            // ставим текстовое значение индикатор на поле усреденных данных 3х моделей - 30
            this.textBox141.Text = "30";
            this.textBox141.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////////////////////////////
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // расчет P(bs) (дБм) при появлении данных в P(bs) (мВт)
            // готовим поля для работы - Коэффициент усиления антенны(Ga) - 0,5
            double val = 10 * Math.Log10(Convert.ToDouble(this.textBox2.Text) / 1);
            this.textBox3.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле - P(bs) (дБм)
            this.textBox3.BackColor = Color.FromArgb(192, 255, 192);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // расчет Длина волны при появлении данных в f
            // готовим поля для работы - Длина волны
            double val = 300000000 / Convert.ToDouble(this.textBox13.Text);
            this.textBox15.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле - Длина волны
            this.textBox15.BackColor = Color.FromArgb(192, 255, 192);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // окно - не даем ему изменять размеры для сохранения удобства 
            // this.WindowState = FormWindowState.Maximized;
            this.MinimumSize = this.Size;
            this.MaximumSize = this.Size;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // обработка модели Хата
            ////////////////////////////////////////////
            ///

            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 10
            // готовим поля для работы
            double val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) * Math.Log10(Convert.ToDouble(this.textBox46.Text))) );
            this.textBox16.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox16.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 10
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) * Math.Log10(Convert.ToDouble(this.textBox46.Text))));
            this.textBox17.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox17.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 10
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) * Math.Log10(Convert.ToDouble(this.textBox46.Text))));
            this.textBox18.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox18.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 10
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) * Math.Log10(Convert.ToDouble(this.textBox46.Text))));
            this.textBox19.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox19.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 10
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) * Math.Log10(Convert.ToDouble(this.textBox46.Text))));
            this.textBox20.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox20.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////
            ///
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 20
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) * Math.Log10(Convert.ToDouble(this.textBox47.Text))));
            this.textBox21.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox21.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 20
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) * Math.Log10(Convert.ToDouble(this.textBox47.Text))));
            this.textBox22.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox22.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 20
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) * Math.Log10(Convert.ToDouble(this.textBox47.Text))));
            this.textBox23.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox23.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 20
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) * Math.Log10(Convert.ToDouble(this.textBox47.Text))));
            this.textBox24.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox24.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 20
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) * Math.Log10(Convert.ToDouble(this.textBox47.Text))));
            this.textBox25.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox25.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 30
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) * Math.Log10(Convert.ToDouble(this.textBox48.Text))));
            this.textBox26.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox26.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 30
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) * Math.Log10(Convert.ToDouble(this.textBox48.Text))));
            this.textBox27.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox27.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 30
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) * Math.Log10(Convert.ToDouble(this.textBox48.Text))));
            this.textBox28.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox28.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 30
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) * Math.Log10(Convert.ToDouble(this.textBox48.Text))));
            this.textBox29.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox29.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 30
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) * Math.Log10(Convert.ToDouble(this.textBox48.Text))));
            this.textBox30.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox30.BackColor = Color.FromArgb(192, 255, 192);


            ////////////////////////////////////////
            /////////////////////////////////////////
            ////////////////////////////////////////
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 40
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) * Math.Log10(Convert.ToDouble(this.textBox49.Text))));
            this.textBox31.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox31.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 40
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) * Math.Log10(Convert.ToDouble(this.textBox49.Text))));
            this.textBox32.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox32.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 40
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) * Math.Log10(Convert.ToDouble(this.textBox49.Text))));
            this.textBox33.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox33.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 40
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) * Math.Log10(Convert.ToDouble(this.textBox49.Text))));
            this.textBox34.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox34.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 40
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) * Math.Log10(Convert.ToDouble(this.textBox49.Text))));
            this.textBox35.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox35.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////
            ///
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 50
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) * Math.Log10(Convert.ToDouble(this.textBox50.Text))));
            this.textBox36.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox36.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 50
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) * Math.Log10(Convert.ToDouble(this.textBox50.Text))));
            this.textBox37.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox37.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 50
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) * Math.Log10(Convert.ToDouble(this.textBox50.Text))));
            this.textBox38.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox38.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 50
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) * Math.Log10(Convert.ToDouble(this.textBox50.Text))));
            this.textBox39.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox39.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 50
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) * Math.Log10(Convert.ToDouble(this.textBox50.Text))));
            this.textBox40.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox40.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 60
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox5.Text)) * Math.Log10(Convert.ToDouble(this.textBox51.Text))));
            this.textBox41.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox41.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 60
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox6.Text)) * Math.Log10(Convert.ToDouble(this.textBox51.Text))));
            this.textBox42.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox42.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 60
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox7.Text)) * Math.Log10(Convert.ToDouble(this.textBox51.Text))));
            this.textBox43.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox43.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 60
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox8.Text)) * Math.Log10(Convert.ToDouble(this.textBox51.Text))));
            this.textBox44.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox44.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 60
            // готовим поля для работы
            val = Convert.ToDouble(this.textBox3.Text) - (46.3 + 33.9 * Math.Log10(Convert.ToDouble(this.textBox13.Text) / 1000000) - 13.8 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) + (44.9 - 6.5 * Math.Log10(Convert.ToDouble(this.textBox9.Text)) * Math.Log10(Convert.ToDouble(this.textBox51.Text))));
            this.textBox45.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox45.BackColor = Color.FromArgb(192, 255, 192);

            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ////////////////////////////////////////////
            ///
            ///// обработка модели Ли
            ////////////////////////////////////////////
            ///
            /// 
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 10
            // готовим поля для работы
            double val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox82.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox52.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox52.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 10
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox82.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox53.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox53.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 10
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox82.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox54.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox54.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 10
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox82.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox55.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox55.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 10
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox82.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox56.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox56.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 20
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox83.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox57.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox57.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 20
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox83.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox58.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox58.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 20
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox83.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox59.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox59.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 20
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox83.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox60.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox60.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 20
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox83.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox61.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox61.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 30
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox84.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox62.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox62.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 30
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox84.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox63.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox63.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 30
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox84.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox64.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox64.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 30
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox84.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox65.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox65.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 30
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox84.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox66.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox66.BackColor = Color.FromArgb(192, 255, 192);
            ///////////////////////////////////
            ///
            ///
            ///
            ///
            ////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 40
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox85.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox67.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox67.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 40
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox85.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox68.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox68.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 40
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox85.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox69.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox69.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 40
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox85.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox70.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox70.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 40
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox85.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox71.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox71.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 50
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox86.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox72.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox72.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 50
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox86.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox73.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox73.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 50
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox86.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox74.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox74.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 50
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox86.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox75.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox75.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 50
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox86.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox76.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox76.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 60
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox87.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox77.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox77.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 60
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox87.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox78.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox78.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 60
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox87.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox79.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox79.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 60
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox87.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox80.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox80.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 60
            // готовим поля для работы
            val = -59 + (Convert.ToDouble(this.textBox3.Text) - 40) - Convert.ToDouble(this.textBox11.Text) * Math.Log10(Convert.ToDouble(this.textBox87.Text) / 1.6) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 30) + Convert.ToDouble(this.textBox4.Text) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3);
            this.textBox81.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox81.BackColor = Color.FromArgb(192, 255, 192);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ////////////////////////////////////////////
            ///
            ///// обработка модели Окамуры
            ////////////////////////////////////////////
            ///
            /// 
            /// 
            /// 
            /// 
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 10
            // готовим поля для работы
            double val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox118.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox124.Text);
            this.textBox88.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox88.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 10
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox118.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox124.Text);
            this.textBox89.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox89.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 10
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox118.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox124.Text);
            this.textBox90.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox90.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 10
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox118.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox124.Text);
            this.textBox91.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox91.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 10
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox118.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox124.Text);
            this.textBox92.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox92.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////

            ////////////////////////////////////////////
            ///
            /// 
            /// 
            /// 
            /// 
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 20
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox119.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox125.Text);
            this.textBox93.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox93.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 20
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox119.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox125.Text);
            this.textBox94.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox94.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 20
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox119.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox125.Text);
            this.textBox95.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox95.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 20
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox119.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox125.Text);
            this.textBox96.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox96.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 20
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox119.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox125.Text);
            this.textBox97.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox97.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////
            ///
            /// 
            /// 
            /// 
            /// 
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 30
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox120.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox126.Text);
            this.textBox98.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox98.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 30
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox120.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox126.Text);
            this.textBox99.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox99.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 30
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox120.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox126.Text);
            this.textBox100.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox100.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 30
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox120.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox126.Text);
            this.textBox101.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox101.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 30
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox120.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox126.Text);
            this.textBox102.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox102.BackColor = Color.FromArgb(192, 255, 192);
            //////////////////////////////////////////////////////////////////////////////////////////
            ///
            //////////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////////////////
            ///
            /// 
            /// 
            /// 
            /// 
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 40
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox121.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox127.Text);
            this.textBox103.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox103.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 40
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox121.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox127.Text);
            this.textBox104.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox104.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 40
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox121.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox127.Text);
            this.textBox105.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox105.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 40
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox121.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox127.Text);
            this.textBox106.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox106.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 40
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox121.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox127.Text);
            this.textBox107.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox107.BackColor = Color.FromArgb(192, 255, 192);

            ////////////////////////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            ///
            /// 
            /// 
            /// 
            /// 
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 50
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox122.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox128.Text);
            this.textBox108.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox108.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 50
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox122.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox128.Text);
            this.textBox109.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox109.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 50
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox122.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox128.Text);
            this.textBox110.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox110.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 50
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox122.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox128.Text);
            this.textBox111.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox111.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 50
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox122.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox128.Text);
            this.textBox112.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox112.BackColor = Color.FromArgb(192, 255, 192);
            /////////////////////////////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            ///
            /// 
            /// 
            /// 
            /// 
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 60
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox123.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox5.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox129.Text);
            this.textBox113.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox113.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 60
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox123.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox6.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox129.Text);
            this.textBox114.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox114.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 60
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox123.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox7.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox129.Text);
            this.textBox115.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox115.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 60
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox123.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox8.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox129.Text);
            this.textBox116.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox116.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 60
            // готовим поля для работы
            val = 10 * Math.Log(Convert.ToDouble(this.textBox2.Text)) + Convert.ToDouble(this.textBox4.Text) - 20 * Math.Log(4 * Math.PI * Convert.ToDouble(this.textBox123.Text) * 1000 / Convert.ToDouble(this.textBox15.Text)) + 20 * Math.Log10(Convert.ToDouble(this.textBox9.Text) / 200) + 10 * Math.Log10(Convert.ToDouble(this.textBox10.Text) / 3) - Convert.ToDouble(this.textBox129.Text);
            this.textBox117.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox117.BackColor = Color.FromArgb(192, 255, 192);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //////////////////////////////////////////////
            //
            // рассчет усредненных данных 3-х моделей
            //
            //////////////////////////////////////////////
            ///
            // Модель Окамуры
            // Уровень сигнала
            // R = 5
            // готовим поля для работы
            double val = ( Convert.ToDouble(this.textBox16.Text) + Convert.ToDouble(this.textBox52.Text) + Convert.ToDouble(this.textBox88.Text) ) / 3;
            this.textBox130.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox130.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала
            // R = 10
            // готовим поля для работы
            val = (Convert.ToDouble(this.textBox21.Text) + Convert.ToDouble(this.textBox57.Text) + Convert.ToDouble(this.textBox93.Text)) / 3;
            this.textBox131.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox131.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала
            // R = 15
            // готовим поля для работы
            val = (Convert.ToDouble(this.textBox26.Text) + Convert.ToDouble(this.textBox62.Text) + Convert.ToDouble(this.textBox98.Text)) / 3;
            this.textBox132.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox132.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала
            // R = 20
            // готовим поля для работы
            val = (Convert.ToDouble(this.textBox31.Text) + Convert.ToDouble(this.textBox67.Text) + Convert.ToDouble(this.textBox103.Text)) / 3;
            this.textBox133.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox133.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала
            // R = 25
            // готовим поля для работы
            val = (Convert.ToDouble(this.textBox36.Text) + Convert.ToDouble(this.textBox72.Text) + Convert.ToDouble(this.textBox108.Text)) / 3;
            this.textBox134.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox134.BackColor = Color.FromArgb(192, 255, 192);

            // Модель Окамуры
            // Уровень сигнала
            // R = 30
            // готовим поля для работы
            val = (Convert.ToDouble(this.textBox41.Text) + Convert.ToDouble(this.textBox77.Text) + Convert.ToDouble(this.textBox113.Text)) / 3;
            this.textBox135.Text = Convert.ToString(val.ToString("F2"));
            // ставим индикатор на поле
            this.textBox135.BackColor = Color.FromArgb(192, 255, 192);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // Выход из программы
            if (Application.MessageLoop)
            {
                // исп.если приложение графическое
                Application.Exit();
            }
            else
            {
                // исп. если приложение консольное
                Environment.Exit(1);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //////////////////////////////////////////////
            //
            // ОЧИСТКА - рассчет усредненных данных 3-х моделей
            //
            //////////////////////////////////////////////
            ///
            // Модель Окамуры
            // Уровень сигнала
            // R = 5
            // очистка
            this.textBox130.Text = "";
            // ставим индикатор на поле
            this.textBox130.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала
            // R = 10
            // очистка
            this.textBox131.Text = "";
            // ставим индикатор на поле
            this.textBox131.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала
            // R = 15
            // очистка
            this.textBox132.Text = "";
            // ставим индикатор на поле
            this.textBox132.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала
            // R = 20
            // очистка
            this.textBox133.Text = "";
            // ставим индикатор на поле
            this.textBox133.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала
            // R = 25
            // очистка
            this.textBox134.Text = "";
            // ставим индикатор на поле
            this.textBox134.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала
            // R = 30
            // очистка
            this.textBox135.Text = "";
            // ставим индикатор на поле
            this.textBox135.BackColor = Color.FromArgb(255, 255, 255);

            ///////////////////////////////////////////////////////////////////////////
            ///

            ////////////////////////////////////////////
            ///
            ///// ОЧИСТКА - модели Окамуры
            ////////////////////////////////////////////
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 10
            // очистка
            this.textBox88.Text = "";
            // ставим индикатор на поле
            this.textBox88.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 10
            // очистка
            this.textBox89.Text = "";
            // ставим индикатор на поле
            this.textBox89.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 10
            // очистка
            this.textBox90.Text = "";
            // ставим индикатор на поле
            this.textBox90.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 10
            // очистка
            this.textBox91.Text = "";
            // ставим индикатор на поле
            this.textBox91.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 10
            // очистка
            this.textBox92.Text = "";
            // ставим индикатор на поле
            this.textBox92.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 20
            // очистка
            this.textBox93.Text = "";
            // ставим индикатор на поле
            this.textBox93.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 20
            // очистка
            this.textBox94.Text = "";
            // ставим индикатор на поле
            this.textBox94.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 20
            // очистка
            this.textBox95.Text = "";
            // ставим индикатор на поле
            this.textBox95.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 20
            // очистка
            this.textBox96.Text = "";
            // ставим индикатор на поле
            this.textBox96.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 20
            // очистка
            this.textBox97.Text = "";
            // ставим индикатор на поле
            this.textBox97.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 30
            // очистка
            this.textBox98.Text = "";
            // ставим индикатор на поле
            this.textBox98.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 30
            // очистка
            this.textBox99.Text = "";
            // ставим индикатор на поле
            this.textBox99.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 30
            // очистка
            this.textBox100.Text = "";
            // ставим индикатор на поле
            this.textBox100.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 30
            // очистка
            this.textBox101.Text = "";
            // ставим индикатор на поле
            this.textBox101.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 30
            // очистка
            this.textBox102.Text = "";
            // ставим индикатор на поле
            this.textBox102.BackColor = Color.FromArgb(255, 255, 255);
            //////////////////////////////////////////////////////////////////////////////////////////
            ///
            //////////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 40
            // очистка
            this.textBox103.Text = "";
            // ставим индикатор на поле
            this.textBox103.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 40
            // очистка
            this.textBox104.Text = "";
            // ставим индикатор на поле
            this.textBox104.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 40
            // очистка
            this.textBox105.Text = "";
            // ставим индикатор на поле
            this.textBox105.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 40
            // очистка
            this.textBox106.Text = "";
            // ставим индикатор на поле
            this.textBox106.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 40
            // очистка
            this.textBox107.Text = "";
            // ставим индикатор на поле
            this.textBox107.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 50
            // очистка
            this.textBox108.Text = "";
            // ставим индикатор на поле
            this.textBox108.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 50
            // очистка
            this.textBox109.Text = "";
            // ставим индикатор на поле
            this.textBox109.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 50
            // очистка
            this.textBox110.Text = "";
            // ставим индикатор на поле
            this.textBox110.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 50
            // очистка
            this.textBox111.Text = "";
            // ставим индикатор на поле
            this.textBox111.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 50
            // очистка
            this.textBox112.Text = "";
            // ставим индикатор на поле
            this.textBox112.BackColor = Color.FromArgb(255, 255, 255);
            /////////////////////////////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Окамуры
            // Уровень сигнала(h=15m)
            // R = 60
            // очистка
            this.textBox113.Text = "";
            // ставим индикатор на поле
            this.textBox113.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=25m)
            // R = 60
            // очистка
            this.textBox114.Text = "";
            // ставим индикатор на поле
            this.textBox114.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=35m)
            // R = 60
            // очистка
            this.textBox115.Text = "";
            // ставим индикатор на поле
            this.textBox115.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=50m)
            // R = 60
            // очистка
            this.textBox116.Text = "";
            // ставим индикатор на поле
            this.textBox116.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Окамуры
            // Уровень сигнала(h=75m)
            // R = 60
            // очистка
            this.textBox117.Text = "";
            // ставим индикатор на поле
            this.textBox117.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            ///
            ///// ОЧИСТКА модели Ли
            ////////////////////////////////////////////
            ///
            /// 
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 10
            // очистка
            this.textBox52.Text = "";
            // ставим индикатор на поле
            this.textBox52.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 10
            // очистка
            this.textBox53.Text = "";
            // ставим индикатор на поле
            this.textBox53.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 10
            // очистка
            this.textBox54.Text = "";
            // ставим индикатор на поле
            this.textBox54.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 10
            // очистка
            this.textBox55.Text = "";
            // ставим индикатор на поле
            this.textBox55.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 10
            // очистка
            this.textBox56.Text = "";
            // ставим индикатор на поле
            this.textBox56.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 20
            // очистка
            this.textBox57.Text = "";
            // ставим индикатор на поле
            this.textBox57.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 20
            // очистка
            this.textBox58.Text = "";
            // ставим индикатор на поле
            this.textBox58.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 20
            // очистка
            this.textBox59.Text = "";
            // ставим индикатор на поле
            this.textBox59.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 20
            // очистка
            this.textBox60.Text = "";
            // ставим индикатор на поле
            this.textBox60.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 20
            // очистка
            this.textBox61.Text = "";
            // ставим индикатор на поле
            this.textBox61.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 30
            // очистка
            this.textBox62.Text = "";
            // ставим индикатор на поле
            this.textBox62.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 30
            // очистка
            this.textBox63.Text = "";
            // ставим индикатор на поле
            this.textBox63.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 30
            // очистка
            this.textBox64.Text = "";
            // ставим индикатор на поле
            this.textBox64.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 30
            // очистка
            this.textBox65.Text = "";
            // ставим индикатор на поле
            this.textBox65.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 30
            // очистка
            this.textBox66.Text = "";
            // ставим индикатор на поле
            this.textBox66.BackColor = Color.FromArgb(255, 255, 255);
            ///////////////////////////////////
            ///
            ////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 40
            // очистка
            this.textBox67.Text = "";
            // ставим индикатор на поле
            this.textBox67.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 40
            // очистка
            this.textBox68.Text = "";
            // ставим индикатор на поле
            this.textBox68.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 40
            // очистка
            this.textBox69.Text = "";
            // ставим индикатор на поле
            this.textBox69.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 40
            // очистка
            this.textBox70.Text = "";
            // ставим индикатор на поле
            this.textBox70.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 40
            // очистка
            this.textBox71.Text = "";
            // ставим индикатор на поле
            this.textBox71.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 50
            // очистка
            this.textBox72.Text = "";
            // ставим индикатор на поле
            this.textBox72.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 50
            // очистка
            this.textBox73.Text = "";
            // ставим индикатор на поле
            this.textBox73.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 50
            // очистка
            this.textBox74.Text = "";
            // ставим индикатор на поле
            this.textBox74.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 50
            // очистка
            this.textBox75.Text = "";
            // ставим индикатор на поле
            this.textBox75.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 50
            // очистка
            this.textBox76.Text = "";
            // ставим индикатор на поле
            this.textBox76.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////

            ////////////////////////////////////////////
            ///
            ////////////////////////////////////////////
            // Модель Ли
            // Уровень сигнала(h=15m)
            // R = 60
            // очистка
            this.textBox77.Text = "";
            // ставим индикатор на поле
            this.textBox77.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=25m)
            // R = 60
            // очистка
            this.textBox78.Text = "";
            // ставим индикатор на поле
            this.textBox78.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=35m)
            // R = 60
            // очистка
            this.textBox79.Text = "";
            // ставим индикатор на поле
            this.textBox79.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=50m)
            // R = 60
            // очистка
            this.textBox80.Text = "";
            // ставим индикатор на поле
            this.textBox80.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Ли
            // Уровень сигнала(h=75m)
            // R = 60
            // очистка
            this.textBox81.Text = "";
            // ставим индикатор на поле
            this.textBox81.BackColor = Color.FromArgb(255, 255, 255);

            //////////////////////////////////////////////////////////////////
            ///
            // ОЧИСТКА модели Хата
            ////////////////////////////////////////////
            ///
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 10
            // очистка
            this.textBox16.Text = "";
            // ставим индикатор на поле
            this.textBox16.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 10
            // очистка
            this.textBox17.Text = "";
            // ставим индикатор на поле
            this.textBox17.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 10
            // очистка
            this.textBox18.Text = "";
            // ставим индикатор на поле
            this.textBox18.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 10
            // очистка
            this.textBox19.Text = "";
            // ставим индикатор на поле
            this.textBox19.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 10
            // очистка
            this.textBox20.Text = "";
            // ставим индикатор на поле
            this.textBox20.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            ///
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 20
            // очистка
            this.textBox21.Text = "";
            // ставим индикатор на поле
            this.textBox21.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 20
            // очистка
            this.textBox22.Text = "";
            // ставим индикатор на поле
            this.textBox22.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 20
            // очистка
            this.textBox23.Text = "";
            // ставим индикатор на поле
            this.textBox23.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 20
            // очистка
            this.textBox24.Text = "";
            // ставим индикатор на поле
            this.textBox24.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 20
            // очистка
            this.textBox25.Text = "";
            // ставим индикатор на поле
            this.textBox25.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 30
            // очистка
            this.textBox26.Text = "";
            // ставим индикатор на поле
            this.textBox26.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 30
            // очистка
            this.textBox27.Text = "";
            // ставим индикатор на поле
            this.textBox27.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 30
            // очистка
            this.textBox28.Text = "";
            // ставим индикатор на поле
            this.textBox28.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 30
            // очистка
            this.textBox29.Text = "";
            // ставим индикатор на поле
            this.textBox29.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 30
            // очистка
            this.textBox30.Text = "";
            // ставим индикатор на поле
            this.textBox30.BackColor = Color.FromArgb(255, 255, 255);


            ////////////////////////////////////////
            /////////////////////////////////////////
            ////////////////////////////////////////
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 40
            // очистка
            this.textBox31.Text = "";
            // ставим индикатор на поле
            this.textBox31.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 40
            // очистка
            this.textBox32.Text = "";
            // ставим индикатор на поле
            this.textBox32.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 40
            // очистка
            this.textBox33.Text = "";
            // ставим индикатор на поле
            this.textBox33.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 40
            // очистка
            this.textBox34.Text = "";
            // ставим индикатор на поле
            this.textBox34.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 40
            // очистка
            this.textBox35.Text = "";
            // ставим индикатор на поле
            this.textBox35.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            ///
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 50
            // очистка
            this.textBox36.Text = "";
            // ставим индикатор на поле
            this.textBox36.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 50
            // очистка
            this.textBox37.Text = "";
            // ставим индикатор на поле
            this.textBox37.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 50
            // очистка
            this.textBox38.Text = "";
            // ставим индикатор на поле
            this.textBox38.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 50
            // очистка
            this.textBox39.Text = "";
            // ставим индикатор на поле
            this.textBox39.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 50
            // очистка
            this.textBox40.Text = "";
            // ставим индикатор на поле
            this.textBox40.BackColor = Color.FromArgb(255, 255, 255);

            ////////////////////////////////////////////
            // Модель Хата
            // Уровень сигнала(h=15m)
            // R = 60
            // очистка
            this.textBox41.Text = "";
            // ставим индикатор на поле
            this.textBox41.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=25m)
            // R = 60
            // очистка
            this.textBox42.Text = "";
            // ставим индикатор на поле
            this.textBox42.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=35m)
            // R = 60
            // очистка
            this.textBox43.Text = "";
            // ставим индикатор на поле
            this.textBox43.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=50m)
            // R = 60
            // очистка
            this.textBox44.Text = "";
            // ставим индикатор на поле
            this.textBox44.BackColor = Color.FromArgb(255, 255, 255);

            // Модель Хата
            // Уровень сигнала(h=75m)
            // R = 60
            // очистка
            this.textBox45.Text = "";
            // ставим индикатор на поле
            this.textBox45.BackColor = Color.FromArgb(255, 255, 255);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            // генерация временной метки для сохрания файла со штампом времени и дня
            string dt_stamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss",CultureInfo.InvariantCulture);

            // создаем рабочую книгу
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Расчёт Окамура Хата Ли");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");

                // Добавить строку
                //List<string[]> headerRow = new List<string[]>()
                //{
                //  new string[] { "ID", "First Name", "Last Name", "DOB" }
                //};

                // определяем диапазон заголовка
                //string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // выбираем активный лист
                var worksheet = excel.Workbook.Worksheets["Расчёт Окамура Хата Ли"];

                // строка данных данного заголовка
                //worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                // определяем заголовки левого края для описаний
                worksheet.Cells["A1"].Value = label1.Text;
                worksheet.Cells["B1"].Value = textBox1.Text;
                //
                worksheet.Cells["A2"].Value = label2.Text;
                worksheet.Cells["B2"].Value = textBox2.Text;
                //
                worksheet.Cells["A3"].Value = label11.Text;
                worksheet.Cells["B3"].Value = textBox3.Text;
                //
                worksheet.Cells["A4"].Value = label3.Text;
                worksheet.Cells["B4"].Value = textBox4.Text;
                //
                worksheet.Cells["A5"].Value = label4.Text;
                worksheet.Cells["B5"].Value = textBox5.Text;
                worksheet.Cells["C5"].Value = textBox6.Text;
                worksheet.Cells["D5"].Value = textBox7.Text;
                worksheet.Cells["E5"].Value = textBox8.Text;
                worksheet.Cells["F5"].Value = textBox9.Text;
                //
                worksheet.Cells["A6"].Value = label5.Text;
                worksheet.Cells["B6"].Value = textBox10.Text;
                //
                worksheet.Cells["A7"].Value = label6.Text;
                worksheet.Cells["B7"].Value = textBox11.Text;
                //
                worksheet.Cells["A8"].Value = label7.Text;
                worksheet.Cells["B8"].Value = textBox12.Text;
                //
                worksheet.Cells["A9"].Value = label8.Text;
                worksheet.Cells["B9"].Value = textBox13.Text;
                //
                worksheet.Cells["A10"].Value = label9.Text;
                worksheet.Cells["B10"].Value = textBox14.Text;

                worksheet.Cells["A11"].Value = label10.Text;
                worksheet.Cells["B11"].Value = textBox15.Text;


                //////////////////////////////////////////////
                ///

                worksheet.Cells["A12"].Value = label42.Text;
                
                worksheet.Cells["B12"].Value = textBox124.Text;
                worksheet.Cells["C12"].Value = textBox125.Text;
                worksheet.Cells["D12"].Value = textBox126.Text;
                worksheet.Cells["E12"].Value = textBox127.Text;
                worksheet.Cells["F12"].Value = textBox128.Text;
                worksheet.Cells["G12"].Value = textBox129.Text;

                ////////////////////////////////////////////////////
                /// МОДЕЛЬ ОКАМУРА
                ///
                worksheet.Cells["A14"].Style.Font.Bold = true;
                worksheet.Cells["A14"].Value = label32.Text;
                worksheet.Cells["A15"].Value = label33.Text;
                worksheet.Cells["A16"].Value = label34.Text;
                worksheet.Cells["A17"].Value = label35.Text;
                worksheet.Cells["A18"].Value = label36.Text;
                worksheet.Cells["A19"].Value = label37.Text;
                worksheet.Cells["A20"].Value = label38.Text;

                worksheet.Cells["B15"].Value = textBox118.Text;
                worksheet.Cells["C15"].Value = textBox119.Text;
                worksheet.Cells["D15"].Value = textBox120.Text;
                worksheet.Cells["E15"].Value = textBox121.Text;
                worksheet.Cells["F15"].Value = textBox122.Text;
                worksheet.Cells["G15"].Value = textBox123.Text;

                ////////////////////////////////////////////////////////////////////////////

                worksheet.Cells["B16"].Value = textBox88.Text;
                worksheet.Cells["C16"].Value = textBox93.Text;
                worksheet.Cells["D16"].Value = textBox98.Text;
                worksheet.Cells["E16"].Value = textBox103.Text;
                worksheet.Cells["F16"].Value = textBox108.Text;
                worksheet.Cells["G16"].Value = textBox113.Text;

                worksheet.Cells["B17"].Value = textBox89.Text;
                worksheet.Cells["C17"].Value = textBox94.Text;
                worksheet.Cells["D17"].Value = textBox99.Text;
                worksheet.Cells["E17"].Value = textBox104.Text;
                worksheet.Cells["F17"].Value = textBox109.Text;
                worksheet.Cells["G17"].Value = textBox114.Text;

                worksheet.Cells["B18"].Value = textBox90.Text;
                worksheet.Cells["C18"].Value = textBox95.Text;
                worksheet.Cells["D18"].Value = textBox100.Text;
                worksheet.Cells["E18"].Value = textBox105.Text;
                worksheet.Cells["F18"].Value = textBox110.Text;
                worksheet.Cells["G18"].Value = textBox115.Text;

                worksheet.Cells["B19"].Value = textBox91.Text;
                worksheet.Cells["C19"].Value = textBox96.Text;
                worksheet.Cells["D19"].Value = textBox101.Text;
                worksheet.Cells["E19"].Value = textBox106.Text;
                worksheet.Cells["F19"].Value = textBox111.Text;
                worksheet.Cells["G19"].Value = textBox116.Text;

                worksheet.Cells["B20"].Value = textBox92.Text;
                worksheet.Cells["C20"].Value = textBox97.Text;
                worksheet.Cells["D20"].Value = textBox102.Text;
                worksheet.Cells["E20"].Value = textBox107.Text;
                worksheet.Cells["F20"].Value = textBox112.Text;
                worksheet.Cells["G20"].Value = textBox117.Text;

                ////////////////////////////////////////////////////
                /// МОДЕЛЬ ХАТА
                worksheet.Cells["A22"].Style.Font.Bold = true;
                worksheet.Cells["A22"].Value = label14.Text;
                worksheet.Cells["A23"].Value = label33.Text;
                worksheet.Cells["A24"].Value = label34.Text;
                worksheet.Cells["A25"].Value = label35.Text;
                worksheet.Cells["A26"].Value = label36.Text;
                worksheet.Cells["A27"].Value = label37.Text;
                worksheet.Cells["A28"].Value = label38.Text;

                worksheet.Cells["B23"].Value = textBox46.Text;
                worksheet.Cells["C23"].Value = textBox47.Text;
                worksheet.Cells["D23"].Value = textBox48.Text;
                worksheet.Cells["E23"].Value = textBox49.Text;
                worksheet.Cells["F23"].Value = textBox50.Text;
                worksheet.Cells["G23"].Value = textBox51.Text;

                ////////////////////////////////////////////////////////////////////////////

                worksheet.Cells["B24"].Value = textBox16.Text;
                worksheet.Cells["C24"].Value = textBox21.Text;
                worksheet.Cells["D24"].Value = textBox26.Text;
                worksheet.Cells["E24"].Value = textBox31.Text;
                worksheet.Cells["F24"].Value = textBox36.Text;
                worksheet.Cells["G24"].Value = textBox41.Text;

                worksheet.Cells["B25"].Value = textBox17.Text;
                worksheet.Cells["C25"].Value = textBox22.Text;
                worksheet.Cells["D25"].Value = textBox27.Text;
                worksheet.Cells["E25"].Value = textBox32.Text;
                worksheet.Cells["F25"].Value = textBox37.Text;
                worksheet.Cells["G25"].Value = textBox42.Text;

                worksheet.Cells["B26"].Value = textBox18.Text;
                worksheet.Cells["C26"].Value = textBox23.Text;
                worksheet.Cells["D26"].Value = textBox28.Text;
                worksheet.Cells["E26"].Value = textBox33.Text;
                worksheet.Cells["F26"].Value = textBox38.Text;
                worksheet.Cells["G26"].Value = textBox43.Text;

                worksheet.Cells["B27"].Value = textBox19.Text;
                worksheet.Cells["C27"].Value = textBox24.Text;
                worksheet.Cells["D27"].Value = textBox29.Text;
                worksheet.Cells["E27"].Value = textBox34.Text;
                worksheet.Cells["F27"].Value = textBox39.Text;
                worksheet.Cells["G27"].Value = textBox44.Text;

                worksheet.Cells["B28"].Value = textBox20.Text;
                worksheet.Cells["C28"].Value = textBox25.Text;
                worksheet.Cells["D28"].Value = textBox30.Text;
                worksheet.Cells["E28"].Value = textBox35.Text;
                worksheet.Cells["F28"].Value = textBox40.Text;
                worksheet.Cells["G28"].Value = textBox45.Text;

                /////////////////////////////////////////////////////////////
                ///
                ////////////////////////////////////////////////////
                /// МОДЕЛЬ ЛИ
                worksheet.Cells["A30"].Style.Font.Bold = true;
                worksheet.Cells["A30"].Value = label21.Text;
                worksheet.Cells["A31"].Value = label33.Text;
                worksheet.Cells["A32"].Value = label34.Text;
                worksheet.Cells["A33"].Value = label35.Text;
                worksheet.Cells["A34"].Value = label36.Text;
                worksheet.Cells["A35"].Value = label37.Text;
                worksheet.Cells["A36"].Value = label38.Text;

                worksheet.Cells["B31"].Value = textBox46.Text;
                worksheet.Cells["C31"].Value = textBox47.Text;
                worksheet.Cells["D31"].Value = textBox48.Text;
                worksheet.Cells["E31"].Value = textBox49.Text;
                worksheet.Cells["F31"].Value = textBox50.Text;
                worksheet.Cells["G31"].Value = textBox51.Text;

                ////////////////////////////////////////////////////////////////////////////

                worksheet.Cells["B32"].Value = textBox52.Text;
                worksheet.Cells["C32"].Value = textBox57.Text;
                worksheet.Cells["D32"].Value = textBox62.Text;
                worksheet.Cells["E32"].Value = textBox67.Text;
                worksheet.Cells["F32"].Value = textBox72.Text;
                worksheet.Cells["G32"].Value = textBox77.Text;

                worksheet.Cells["B33"].Value = textBox53.Text;
                worksheet.Cells["C33"].Value = textBox58.Text;
                worksheet.Cells["D33"].Value = textBox63.Text;
                worksheet.Cells["E33"].Value = textBox68.Text;
                worksheet.Cells["F33"].Value = textBox73.Text;
                worksheet.Cells["G33"].Value = textBox78.Text;

                worksheet.Cells["B34"].Value = textBox54.Text;
                worksheet.Cells["C34"].Value = textBox59.Text;
                worksheet.Cells["D34"].Value = textBox64.Text;
                worksheet.Cells["E34"].Value = textBox69.Text;
                worksheet.Cells["F34"].Value = textBox74.Text;
                worksheet.Cells["G34"].Value = textBox79.Text;

                worksheet.Cells["B35"].Value = textBox55.Text;
                worksheet.Cells["C35"].Value = textBox60.Text;
                worksheet.Cells["D35"].Value = textBox65.Text;
                worksheet.Cells["E35"].Value = textBox70.Text;
                worksheet.Cells["F35"].Value = textBox75.Text;
                worksheet.Cells["G35"].Value = textBox80.Text;

                worksheet.Cells["B36"].Value = textBox56.Text;
                worksheet.Cells["C36"].Value = textBox61.Text;
                worksheet.Cells["D36"].Value = textBox66.Text;
                worksheet.Cells["E36"].Value = textBox71.Text;
                worksheet.Cells["F36"].Value = textBox76.Text;
                worksheet.Cells["G36"].Value = textBox81.Text;

                /////////////////////////////////////////////////////////////
                ///

                /////////////////////////////////////////////////////////////
                ///
                ////////////////////////////////////////////////////
                /// УСРЕДНЕННЫЕ ДАННЫЕ ДЛЯ 3-х МОДЕЛЕЙ
                worksheet.Cells["A38"].Style.Font.Bold = true;
                worksheet.Cells["A38"].Value = "Усреднённые данные 3-х моделей";
                worksheet.Cells["A39"].Value = label47.Text;
                worksheet.Cells["A40"].Value = label48.Text;

                worksheet.Cells["B39"].Value = textBox136.Text;
                worksheet.Cells["C39"].Value = textBox137.Text;
                worksheet.Cells["D39"].Value = textBox138.Text;
                worksheet.Cells["E39"].Value = textBox139.Text;
                worksheet.Cells["F39"].Value = textBox140.Text;
                worksheet.Cells["G39"].Value = textBox141.Text;

                ////////////////////////////////////////////////////////////////////////////

                worksheet.Cells["B40"].Value = textBox130.Text;
                worksheet.Cells["C40"].Value = textBox131.Text;
                worksheet.Cells["D40"].Value = textBox132.Text;
                worksheet.Cells["E40"].Value = textBox133.Text;
                worksheet.Cells["F40"].Value = textBox134.Text;
                worksheet.Cells["G40"].Value = textBox135.Text;

                /////////////////////////////////////////////////////////////
                ///





                // определяем стиль
                worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A2"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A3"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A4"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A5"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A6"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A7"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A8"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A9"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A10"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A11"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                worksheet.Cells["A12"].Style.Font.Color.SetColor(System.Drawing.Color.Navy);

                //////////////////////////////////////////////
                // определяем заголовки для данных
                //worksheet.Cells["B1"].Value = textBox2.Text;
                //worksheet.Cells["B2"].Value = textBox1.Text;
                //
                // определяем стиль

                /////////////////////////////////////////////////////////////////////////////////
                ///
                // опеределяем имя файла и путь
                FileInfo excelFile = new FileInfo(@Directory.GetCurrentDirectory() + "\\Экспорт_данных\\" + "Расчёт Окамура Хата Ли_" + dt_stamp + ".xlsx");

                // сохраняем
                excel.SaveAs(excelFile);

                MessageBox.Show("Файл {Расчёт Окамура Хата Ли_" + dt_stamp + ".xlsx" + "}, сохранен в каталоге экспорта по адресу - " + Directory.GetCurrentDirectory() + "\\Экспорт_данных\\", "Экспорт данных был произведен..", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
    }
}
