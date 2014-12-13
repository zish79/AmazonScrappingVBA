using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Reflection;




namespace AmazaonScrapping
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int ioexception = 0;
            int webexception = 0;
            string MyFile = @"d:\"+textBox4.Text.ToString().Trim()+" "+textBox1.Text.ToString().Trim()+"-"+textBox3.Text.ToString().Trim()+".xls";
            Microsoft.Office.Interop.Excel.Application xl = null;
            Microsoft.Office.Interop.Excel._Workbook wb = null;
            Microsoft.Office.Interop.Excel._Worksheet sheet = null;
            //VBIDE.VBComponent module = null;
            bool SaveChanges = false;

            if (File.Exists(MyFile)) { File.Delete(MyFile); }


            xl = new Microsoft.Office.Interop.Excel.Application();
            xl.Visible = false;

            wb = (Microsoft.Office.Interop.Excel._Workbook)(xl.Workbooks.Add(Missing.Value));
            wb.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            wb.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            wb.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            wb.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            wb.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            sheet = (Microsoft.Office.Interop.Excel._Worksheet)(wb.Sheets[1]);

            //for (int r = 0; r < 20; r++)
            //{

            //    for (int c = 0; c < 10; c++)
            //    {
            //        sheet.Cells[r + 1, c + 1] = 125;
            //    }
            //}





            // Let loose control of the Excel instance



            // Get the stream from the returned web response
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            string strLine;
            // Read the stream a line at a time and place each one
            // into the stringbuilder
            int Pr_Count = 0;
            int r = 2;
            int c = 0;
            int starter = 0;
            int Ender = 0;
            try
            {
                 starter = Convert.ToInt32(textBox1.Text.ToString());
                 Ender = Convert.ToInt32(textBox3.Text.ToString());
            }
            catch(Exception ex)
            {
                MessageBox.Show("Enter Valid Page Number");
                return;
            }
            if (!string.IsNullOrEmpty(textBox2.Text.ToString().Trim()))
            {


                for (int ii = starter; ii <= Ender; ii++)
                {
                    Pr_Count++;
                    if (Pr_Count >= 25)
                    {
                        MessageBox.Show("Buy A Premium Tool in Just 50$ to Fetch 100% Data");
                        return;
                    }
                    sheet.Cells[r, c + 1] = "Page: "+ii.ToString();
                    r++;
                    string url1 = "http://www.amazon.com/gp/shops/storefront/index.html?ie=UTF8&isSearch=1&page=";
                    string url2 = "&searchTerm=&sellerID=" + textBox2.Text.ToString().Trim() + "&sortBy=StartDateDesc";
                    try
                    {
                        WebRequest req = WebRequest.Create(url1 + ii + url2);
                        if (req != null)
                        {


                            if (req.GetResponse() != null)
                            {
                                // Get the stream from the returned web response

                                StreamReader stream = new StreamReader(req.GetResponse().GetResponseStream());
                                if (stream != null)
                                {
                                    while ((strLine = stream.ReadLine()) != null)
                                    {
                                        //// Ignore blank lines
                                        //if (strLine.Length > 0)
                                        //    sb.Append(strLine);
                                        string sss = "";
                                        if (strLine.Contains("<span class='small'><a href="))
                                        {
                                            int last = strLine.IndexOf("<span class='small'><a href=");
                                            if (last != -1)
                                            {
                                                last = last + 41;
                                                char[] array = strLine.ToCharArray();
                                                while (array[last] != '/')
                                                {
                                                    sss += array[last];

                                                    last++;
                                                }
                                                sheet.Cells[r + 1, c + 1] = sss;
                                                r++;
                                            }
                                        }

                                        //if (strLine.Contains("class=h3color"))
                                        //{
                                        //    if (strLine.Contains("No results found."))
                                        //    {
                                        //        MessageBox.Show("Enter Correct Seller ID");
                                        //        return;
                                        //    }
                                            
                                        //}

                                        sss = "";
                                        //if (strLine.Contains("<br>Buy New: <a href="))
                                        //{
                                        //    char[] array = strLine.ToCharArray();
                                        //    int size = array.Length;
                                        //    for (int i = 0; i < size - 1; i++)
                                        //    {
                                        //        if (array[i] == '>' && array[i + 1] == '$')
                                        //        {
                                        //            int k = i;
                                        //            for (int j = k + 1; array[j] != '<'; j++, i++)
                                        //            {
                                        //                sb.Append(array[j]);

                                        //                sss += array[j];

                                        //            }
                                        //            sheet.Cells[r + 1, c + 6] = sss;
                                        //        }
                                        //    }
                                        //    r++;
                                        //}

                                        //if (strLine.Contains("Currently Unavailable"))
                                        //{

                                        //    sheet.Cells[r + 1, c + 6] = "Currently Unavailable";
                                        //    r++;
                                        //}

                                    }


                                }
                                stream.Close();
                            }

                        }
                    }
                    catch (WebException)
                    {
                        webexception++;
                        MessageBox.Show("Unable to Fetch Data from Server Against Page Number: "+ii.ToString()+"\n");
                        MessageBox.Show("Sorry! Execution Stops Itself at this Page Number");
                        sheet.Name = "Movie";
                        sheet.Cells[1, 1] = "ASIN";
                        xl.Visible = false;
                        xl.UserControl = false;
                        SaveChanges = true;

                        wb.SaveAs(MyFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                                  null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                                  false, false, null, null, null);
                        xl.Visible = false;
                        xl.UserControl = false;
                        // Close the document and avoid user prompts to save if our method failed.
                        wb.Close(SaveChanges, null, null);
                        xl.Workbooks.Close();
                        Dispose();
                    }
                    catch (IOException)
                    {
                        ioexception++;
                        MessageBox.Show("Unable to Fetch Data from Server Against Page Number: "+ii.ToString()+"\n");
                        MessageBox.Show("Sorry! Execution Stops Itself at this Page Number");
                        sheet.Name = "Movie";
                        sheet.Cells[1, 1] = "ASIN";

                        xl.Visible = false;
                        xl.UserControl = false;

                        SaveChanges = true;


                        wb.SaveAs(MyFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                                  null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                                  false, false, null, null, null);
                        xl.Visible = false;
                        xl.UserControl = false;
                        // Close the document and avoid user prompts to save if our method failed.
                        wb.Close(SaveChanges, null, null);
                        xl.Workbooks.Close();
                        Dispose();
                    }
                    catch (Exception)
                    {
                        ioexception++;
                        MessageBox.Show("Unable to Fetch Data from Server Against Page Number: " + ii.ToString() + "\n");
                        MessageBox.Show("Sorry! Execution Stops Itself at this Page Number");
                        sheet.Name = "Movie";
                        sheet.Cells[1, 1] = "ASIN";

                        xl.Visible = false;
                        xl.UserControl = false;

                        SaveChanges = true;


                        wb.SaveAs(MyFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                                  null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                                  false, false, null, null, null);
                        xl.Visible = false;
                        xl.UserControl = false;
                        // Close the document and avoid user prompts to save if our method failed.
                        wb.Close(SaveChanges, null, null);
                        xl.Workbooks.Close();
                        Dispose();
                    }
                    r = r + 2;
                    

                }
            }

            // Finished with the stream so close it now



            // set come column heading names
            sheet.Name = "Movie";
            sheet.Cells[1, 1] = "ASIN";

            xl.Visible = false;
            xl.UserControl = false;

            // Set a flag saying that all is well and it is ok to save our changes to a file.

            SaveChanges = true;

            //  Save the file to disk

            wb.SaveAs(MyFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                      null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                      false, false, null, null, null);
            xl.Visible = false;
            xl.UserControl = false;
            // Close the document and avoid user prompts to save if our method failed.
            wb.Close(SaveChanges, null, null);
            xl.Workbooks.Close();
            Dispose();
        }
    }
}
