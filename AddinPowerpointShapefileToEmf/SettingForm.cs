using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using System.Net;
using System.Collections.Specialized;

namespace AddinPowerpointShapefileToEmf
{
    public partial class SettingForm : Form
    {
        OpenFileDialog openfile;
        string nomFichier;
        List<string> listShapes;
        List<string> listShapesCoordinates;
        public SettingForm()
        {
            InitializeComponent();
            listShapes = new List<string>();

            // Add a link to the LinkLabel.
            
        }

        private void buttonOpenFile_Click(object sender, EventArgs e)
        {
            openfile = new OpenFileDialog();
            try
            {
                openfile.ShowDialog();
                nomFichier = openfile.FileName;

                if (File.Exists(nomFichier))
                {
                    string[] fileInfo = nomFichier.Split('.');
                    string extension = fileInfo[fileInfo.Length-1];

                    if (extension.Equals("txt"))
                    {
                        string appendText = "";

                        appendText = File.ReadAllText(nomFichier);
                        string[] arrayFileContent = appendText.Split('#');

                        for (int i = 0; i < arrayFileContent.Length; i++)
                        {
                            if (!arrayFileContent[i].Trim().Equals(""))
                            {
                                listShapes.Add(arrayFileContent[i]);
                            }
                        }

                        if (listShapes.Count != 0)
                        {
                            //buttonCreate.Enabled = true;
                            CreateShapes();
                            //MessageBox.Show(listShapes.Count + " shapes are detected in the file");
                        }
                        else
                        {
                            MessageBox.Show("No shape detected in the file, please verify the file or contact the support");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select a txt file converted on the web application, the name should be (ShapeFileTo.Emf ...)");
                        //buttonCreate.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error has occured close and try again or take a screenshot and send to mane2@un.org : "+ex.Message);
               // buttonCreate.Enabled = false;
            }
        }

        private void CreateShapes()
        {

            

            


            int lastPoint = 0;

            try
            {
                WriteLog("emf", "Start");
                if (listShapes.Count != 0)
                {
                    WriteLog("emf", " - nb shapes : " + listShapes.Count);
                    Slides listSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                    
                    foreach (Slide maSlide in listSlide)
                    {
                        

                        ShapeRange srd = maSlide.Shapes.Range();
                        srd.Delete();
                        WriteLog("emf", " - all existing slides deleted");

                        List<object> listNames = new List<object>();
                        int c = 0;
                        foreach (string shapefile in listShapes)
                        {
                            string[] shapeInfo = shapefile.Split('*');
                            string shapeName = shapeInfo[0];
                            string[] polygons = shapeInfo[1].Split('_');

                            WriteLog("emf", " - Working on shape : " + shapeName + " - " + polygons.Length + " polygones");

                            for (int i = 0; i < polygons.Length; i++)
                            {
                                if (polygons[i] != "")
                                {
                                    string[] coordinates = polygons[i].Split(',');
                                    if (coordinates.Length > 1)
                                    {
                                        /*
                                        
                                        
                                        for (int j = 0; j < (coordinates.Length - 1); j++)
                                        {
                                            float val = (float)Convert.ToDouble(coordinates[j].Replace('.', ','));

                                            if (y == 0)
                                            {
                                                
                                                points[x, y] = (val);
                                            }
                                            if (y == 1)
                                            {
                                                points[x, y] = (val);
                                            }
                                            x2 = x;

                                            y++;
                                            if (y == 2)
                                            {
                                                y = 0;
                                                x++;
                                            }

                                            lastPoint++;
                                        }
                                        
                                        */

                                        //Single[,] myPoints = new Single[5, 2]; // {500,50,505,55,510,60,515,65,520,70};
                                        int nbPoints = (coordinates.Length - 1) / 2;
                                        WriteLog("emf", "    " + nbPoints + " points to create");
                                        Single[,] points = new Single[nbPoints, 2];                                       
                                        int x = 0;
                                        int y = 0;
                                        int x2 = 0;

                                        for (int j = 0; j < (coordinates.Length - 1); j++)
                                        {
                                            double val = double.Parse(coordinates[j], System.Globalization.CultureInfo.InvariantCulture);

                                            if (y == 0)
                                            {
                                                points[x, y] = (float)val;
                                            }
                                            if (y == 1)
                                            {
                                                points[x, y] = (float)val;
                                            }
                                            x2 = x;

                                            y++;
                                            if (y == 2)
                                            {
                                                y = 0;
                                                x++;
                                            }

                                            lastPoint++;
                                        }

                                   

                                        object po = points;
                                       // WriteLog("emf", "    Shape " + shapeName + " ready to be drawed with " + ((lastPoint) / 2) + "/" + nbPoints + " points, x=" + x + " , y=" + y + ", x2=" + x2);
                                        lastPoint = 0;
                                        maSlide.Shapes.AddPolyline(po).Name = shapeName;
                                        WriteLog("emf", "    Shape " + shapeName + " drawed");
                                        listNames.Add(shapeName);

                                    }
                                }
                            }
                        }

                        object[] arrayNames = new object[listNames.Count];

                        foreach (object name in listNames)
                        {
                            arrayNames[c] = name;
                            c++;
                        }
                        ShapeRange sr = maSlide.Shapes.Range();

                        WriteLog("emf", " - group all shapes");
                        sr.Group().Name = "MonGoupe";
                        sr.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipVertical);
                        

                        WriteLog("emf", " - prepare to resize the shapes");
                        float longeur = sr.Width;
                        float hauteur = sr.Height;

                        int maxWidth = 1280;
                        int maxHeight = 500;

                        float coef = 0;

                        if (longeur < hauteur)
                        {
                            c = maxWidth / Convert.ToInt32(longeur);
                            longeur = longeur * c;
                            hauteur = hauteur * c;
                        }
                        else
                        {
                            c = maxHeight / Convert.ToInt32(hauteur);
                            longeur = longeur * c;
                            hauteur = hauteur * c;
                        }

                        sr.Width = longeur;
                        sr.Height = hauteur;
                        WriteLog("emf", " - Shapes reseized");
                        WriteLog("emf", " - Process End");
                        WriteLog("emf", " - ");
                        WriteLog("emf", " - ");
                        
                    }

                   // MessageBox.Show("Finished");
                    this.Close();

                    //Set myDocument = ActivePresentation.Slides(1)
                    //With myDocument.Shapes
                    //    .AddShape(msoShapeCan, 50, 10, 100, 200).Name = "shpOne"
                    //    .AddShape(msoShapeCube, 150, 250, 100, 200).Name = "shpTwo"
                    //    With.Range(Array("shpOne", "shpTwo")).Group
                    //    End With
                    //End With


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"An error occurred while creating the map, please send the log files in the folder 'C:\OCHA_logs' to the support mane2@un.org");
                WriteLog("emf", " - Error Message : " + ex.Message);
                WriteLog("emf", " - Error StackTrace : " + ex.StackTrace);
                WriteLog("emf", " - Error TargetSite : " + ex.TargetSite);
                WriteLog("emf", " - Error InnerException : " + ex.InnerException);
                WriteLog("emf", " - Error ToString : " + ex.ToString());
                WriteLog("emf", " - Error GetBaseException : " + ex.GetBaseException().ToString());
                
            }
            finally
            {

            }
        }
        

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            
        }

        public void WriteLog(String fichier, String message)
        {
            try
            {
                if (!Directory.Exists(@"C:\OCHA_logs"))
                {
                    Directory.CreateDirectory(@"C:\OCHA_logs");
                }


                String CheminFichier =  @"C:\OCHA_logs\" + fichier + " " + DateTime.Now.ToString("yyyy MM dd") + ".txt";
                if (!System.IO.File.Exists(CheminFichier))
                {
                    System.IO.FileStream fichierLog = System.IO.File.Create(CheminFichier);
                    fichierLog.Close();
                }
                File.AppendAllText(CheminFichier, DateTime.Now + " : " + message + "\r\n");
               
            }
            catch (Exception ex)
            {

            }
        }



        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("mailto:mane2@un.org");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Link.LinkData as string);
        }

        private void SettingForm_Load(object sender, EventArgs e)
        {
            LinkLabel.Link link = new LinkLabel.Link();
            link.LinkData = "https://github.com/chamsocha/Shapefile-to-Emf/releases";
            //linkLabelGitHUb.Links.Add(link);
        }

        private void linkLabelGitHUb_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel.Link link = new LinkLabel.Link();
            link.LinkData = "https://github.com/chamsocha/Shapefile-to-Emf/releases";
            e.Link.LinkData = link.LinkData;
            System.Diagnostics.Process.Start(e.Link.LinkData as string);
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
