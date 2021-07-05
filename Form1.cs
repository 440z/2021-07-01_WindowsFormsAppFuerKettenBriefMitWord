using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace _2021_07_01_WindowsFormsAppFuerKettenBriefMitWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Heading 1";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();

            //Insert a paragraph at the end of the document.
            Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "Heading 2";
            oPara2.Format.SpaceAfter = 6;
            oPara2.Range.InsertParagraphAfter();

            //Insert another paragraph.
            Word.Paragraph oPara3;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:";
            oPara3.Range.Font.Bold = 0;
            //oPara3.Format.SpaceAfter = 24;
            oPara3.Range.InsertParagraphAfter();








            //Insert my own paragraph
            Word.Paragraph oPara42;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara42 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara42.Range.Text = "Insert my own paragraph.      " +
                "An other line of text" +
                "\n" +
                "Wie macht man einen Zeilenumbruch der nicht 3cm hoch ist?"; // ??? Wie macht man einen Zeilenumbruch?
            //oPara42.Range.Text = "An other line of text";
            oPara42.Range.InsertParagraphAfter();


            //Insert my additional own paragraph
            //string string1 = "foo";
            string string1 = textBox1.Text;
            string string2 = textBox2.Text;
            Word.Paragraph oPara43;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara43 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara43.Range.Text = "Exactly here: " + string1 + "\n" +
                "and here: " + string2 + "\n" +
                "are my strings";
            oPara43.Format.SpaceAfter = 24;
            oPara43.Range.InsertParagraphAfter();

            









            //Insert a 3 x 5 table, fill it with data, and make the first row
            //bold and italic.
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 3, 5, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            int r, c;
            string strText;
            for (r = 1; r <= 3; r++)
                for (c = 1; c <= 5; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;

            //Add some text after the table.
            Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "And here's another table:";
            oPara4.Format.SpaceAfter = 24;
            oPara4.Range.InsertParagraphAfter();


            //Insert a 5 x 2 table, fill it with data, and change the column widths.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            for (r = 1; r <= 5; r++)
                for (c = 1; c <= 2; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Columns[1].Width = oWord.InchesToPoints(2); //Change width of columns 1 & 2
            oTable.Columns[2].Width = oWord.InchesToPoints(3);

            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            object oPos;
            //double dPos = oWord.InchesToPoints(7);
            double dPos = oWord.InchesToPoints(10);
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            do
            {
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.ParagraphFormat.SpaceAfter = 6;
                wrdRng.InsertAfter("A line of text");
                wrdRng.InsertParagraphAfter();
                oPos = wrdRng.get_Information
                                       (Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            while (dPos >= Convert.ToDouble(oPos));


            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Word.WdBreakType.wdPageBreak;
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertBreak(ref oPageBreak);
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            wrdRng.InsertParagraphAfter();

            ////Insert a chart.
            //Word.InlineShape oShape;
            //object oClassType = "MSGraph.Chart.8";
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
            //ref oMissing, ref oMissing, ref oMissing,
            //ref oMissing, ref oMissing, ref oMissing);

            ////Demonstrate use of late bound oChart and oChartApp objects to
            ////manipulate the chart object with MSGraph.
            //object oChart;
            //object oChartApp;
            //oChart = oShape.OLEFormat.Object;
            //oChartApp = oChart.GetType().InvokeMember("Application",
            //BindingFlags.GetProperty, null, oChart, null);

            ////Change the chart type to Line.
            //object[] Parameters = new Object[1];
            //Parameters[0] = 4; //xlLine = 4
            //oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
            //null, oChart, Parameters);

            ////Update the chart image and quit MSGraph.
            //oChartApp.GetType().InvokeMember("Update",
            //BindingFlags.InvokeMethod, null, oChartApp, null);
            //oChartApp.GetType().InvokeMember("Quit",
            //BindingFlags.InvokeMethod, null, oChartApp, null);
            ////... If desired, you can proceed from here using the Microsoft Graph 
            ////Object model on the oChart and oChartApp objects to make additional
            ////changes to the chart.

            ////Set the width of the chart.
            //oShape.Width = oWord.InchesToPoints(6.25f);
            //oShape.Height = oWord.InchesToPoints(3.57f);

            ////Add text after the chart.
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //wrdRng.InsertParagraphAfter();
            //wrdRng.InsertAfter("THE END.");

            //Close this form.
            this.Close();



            ////object oTemplate = "c:\\MyTemplate.dot";

            ////Dr. Heuer
            ////object oTemplate = @"C:\Users\ITA8-TN04\OneDrive - IT-Akademie Dr. Heuer GmbH\Praktikum\MyTemplate.docx";

            ////Zu Hause
            ////object oTemplate = @"C:\Users\Windows10\OneDrive - IT-Akademie Dr. Heuer GmbH\Praktikum\MyTemplate.docx";
            //object oTemplate = @"C:\Users\Windows10\OneDrive - IT-Akademie Dr. Heuer GmbH\Formular aus meinem Ordner.docx";


            //oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
            //ref oMissing, ref oMissing);

            //object oBookMark = "MyBookmark";
            ////oDoc.Bookmarks.Item(ref oBookMark).Range.Text = "Some Text Here";
            ////oDoc.Bookmarks[ref oBookMark].Range.Text = "Some Text Here";
            ////oDoc.Bookmarks[ref oBookMark].Range.Text = "endofdoc";
            ////oDoc.Bookmarks[ref oBookMark].Range.Text = "What the curse!!!";

            

            ////oDoc.Bookmarks[oBookMark].Range.Text = "Some Text Here";
            //// Stackoverflow: Frage aus Forum


            //object oStyleName = "MyStyle";
            ////oDoc.Bookmarks.Item(ref oBookMark).Range.set_Style(ref oStyleName);
            ////oDoc.Bookmarks[ref oBookMark].Range.set_Style(ref oStyleName);


            ////object oStyleName = "MyStyle";
            ////oWord.Selection.set_Style(ref oStyleName);





            //// Delete old stuff



            ////Insert my additional own paragraph
            //string string1 = "foo";
            //string string2 = "bar";
            //Word.Paragraph oPara43;
            //oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oPara43 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara43.Range.Text = "Exactly here: " + string1 + "\n" +
            //    "and here: " + string2 + "\n" +
            //    "are my strings"; 
            //oPara43.Range.InsertParagraphAfter();

        }





        

        


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load_1(object sender, EventArgs e)
        {
            //MessageBox.Show("program is running!!");
            


        }

        private void button2_Click(object sender, EventArgs e)
        {
            string inhaltTextbox;
            inhaltTextbox = textBox1.Text;
            MessageBox.Show(inhaltTextbox);
            
        }

        

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
       
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string inhaltTextbox;
            inhaltTextbox = textBox2.Text;
            MessageBox.Show(inhaltTextbox);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
