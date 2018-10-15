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
using RestSharp;

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
                            buttonCreate.Enabled = true;
                            MessageBox.Show(listShapes.Count + " shapes are detected in the file");
                        }
                        else
                        {
                            buttonCreate.Enabled = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select a txt file converted on the web application, the name should be (ShapeFileTo.Emf ...)");
                        buttonCreate.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error has occured close and try again or take a screenshot and send to mane2@un.org : "+ex.Message);
                buttonCreate.Enabled = false;
            }
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            if (listShapes.Count != 0)
            {
                Slides listSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                
                foreach (Slide maSlide in listSlide)
                {
                    ShapeRange srd = maSlide.Shapes.Range();
                    srd.Delete();

                    List<object> listNames = new List<object>();
                    int c = 0;
                    foreach (string shapefile in listShapes)
                    {
                        string[] shapeInfo = shapefile.Split('*');
                        string shapeName = shapeInfo[0];
                        string[] polygons = shapeInfo[1].Split('_');
                        for (int i = 0; i < polygons.Length; i++)
                        {
                            string[] coordinates = polygons[i].Split(',');
                            if (coordinates.Length>1)
                            {
                                Single[,] points = new Single[(coordinates.Length - 1) / 2, 2];
                                int x = 0;
                                int y = 0;
                                for (int j = 0; j < coordinates.Length - 1; j++)
                                {
                                    float val = (float)Convert.ToDouble(coordinates[j].Replace('.', ','));
                                    int zoom = 2;
                                    if (y == 0)
                                    {
                                        
                                        points[x, y] = (val);
                                    }
                                    if (y == 1)
                                    {
                                        points[x, y] = (val);
                                    }

                                    y++;
                                    if (y == 2)
                                    {
                                        y = 0;
                                        x++;
                                    }
                                }

                                object po = points;

                                maSlide.Shapes.AddPolyline(po).Name = shapeName;
                                listNames.Add(shapeName);
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
                    
                    sr.Group().Name ="MonGoupe";
                    sr.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipVertical);
                    float longeur = sr.Width;
                    float hauteur = sr.Height;

                    int maxWidth = 1280;
                    int maxHeight = 500;

                    float coef = 0;

                    if (longeur < hauteur)
                    {
                        c =  maxWidth / Convert.ToInt32(longeur);
                        longeur = longeur * c;
                        hauteur = hauteur * c;
                    }else
                    {
                        c = maxHeight / Convert.ToInt32(hauteur);
                        longeur = longeur * c;
                        hauteur = hauteur * c;
                    }

                    sr.Width = longeur;
                    sr.Height = hauteur;
                }

                MessageBox.Show("Finished");
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

        

        private void button1_Click(object sender, EventArgs e)
        {
            var client = new RestClient("https://ogre.adc4gis.com/convert");
            // client.Authenticator = new HttpBasicAuthenticator(username, password);

            var request = new RestRequest(Method.POST);

            request.AddHeader("Content-Type", "multipart/mixed");

            // add files to upload (works with compatible verbs)
            request.AddFile("file", @"C:\test.zip");

            // execute the request
            client.ExecuteAsync(request, response => {
                Console.WriteLine(response);
            });
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
    }
}
