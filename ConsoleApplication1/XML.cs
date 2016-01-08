//----------------------------------------------------------------------------------------------
// <copyright file="XML.cs" company="Almaviva TSF" author="Gabriele Mazzocca  g.mazzocca@almaviva.it">
// Copyright (c) Almaviva TSF.  All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------
namespace ConsoleAppTFS
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Management;
    using System.Security.AccessControl;
    using System.Text;
    using System.Xml;

    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.VersionControl.Client;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    using AlmaLogger;

    /// <summary>
    /// gestione file xml
    /// </summary>
    public class XML
    {
        /// <summary>
        /// analizza i dati ricavandone i Function point
        /// </summary>
        /// <param name="linguaggio">
        /// lista linguaggi da analizzare
        /// </param>
        /// <param name="fp">
        /// lista di dati relativo ai FP da ricavare
        /// </param>
        /// <returns>
        /// FP totali
        /// </returns>
        public static double AnalisiFP(List<FP> linguaggio, List<FP> fp)
        {
            double sum = 0;
            bool trovato = false;
            foreach (FP item in linguaggio)
            {
                trovato = false;
                foreach (FP tmp in fp)
                {
                    if (item.Linguaggio.ToUpper().Equals(tmp.Linguaggio))
                    {
                        trovato = true;
                        sum += item.Dato / tmp.Dato;
                    }
                }

                if (!trovato)
                {
                    sum += item.Dato / 43;
                }
            }

            return sum;
        }

        /// <summary>
        /// appende TP ad un file .xls leggendo i dati del report dal file xml del cloc
        /// </summary>
        /// <param name="tp">
        /// tp di riferimento
        /// </param>
        /// <param name="collection">
        /// collection di riferimento
        /// </param>
        /// <param name="filename">
        /// excel a cui appendere i file..
        /// </param>
        /// <param name="cIS">
        /// Configuration item Software
        /// </param>
        public static void AppendTP(string tp, string collection, string filename, string cIS)
        {
            XmlDocument xd = new XmlDocument();

            // create new instance of XmlDocument
            XmlDocument doc = new XmlDocument();

            // load from file
            if (File.Exists(filename))
            {
                doc.Load(filename);

                // create node and add value
                XmlNode node = doc.CreateNode(XmlNodeType.Element, "TeamProject", null);
                node.InnerText = tp;
                XmlNode node2 = doc.CreateNode(XmlNodeType.Element, "Collection", null);
                node2.InnerText = collection;
                XmlNode node3 = doc.CreateNode(XmlNodeType.Element, "CISoftware", null);
                node3.InnerText = cIS;

                // add to elements collection
                doc.DocumentElement.FirstChild.AppendChild(node);
                doc.DocumentElement.FirstChild.AppendChild(node2);
                doc.DocumentElement.FirstChild.AppendChild(node3);

                // save back
                doc.Save(filename);
            }
        }

        /// <summary>
        /// metodo di controllo dei totali(FP e Costi)
        /// </summary>
        public static void ControlloTotali()
        {
            string repDir = System.Configuration.ConfigurationManager.AppSettings["repDir"];
            DirectoryInfo dinfo = Directory.CreateDirectory(repDir);
            List<string> pass = new List<string>();

            var elenco = dinfo.EnumerateFiles("*.xml", SearchOption.AllDirectories);
            foreach (FileInfo f in elenco)
            {
                using (XmlTextReader reader = new XmlTextReader(f.FullName))
                {
                    while (reader.Read())
                    {
                        if (reader.Name.Equals("language") && reader.NodeType.Equals(XmlNodeType.Element))
                        {
                            int inizio = f.Name.IndexOf("#") + 1;
                            int fine = f.Name.Length - 4;
                            pass.Add(f.Name.Substring(inizio, fine - inizio));
                            while (reader.MoveToNextAttribute())
                            {
                                if (reader.Name.Equals("name") || reader.Name.Equals("files_count") || reader.Name.Equals("code"))
                                {
                                    pass.Add(reader.Value);
                                }
                            }
                        }
                    }
                }

                double totLine = 0;
                List<FP> fp = XML.RecuperaFP();
                double costi = double.Parse(System.Configuration.ConfigurationManager.AppSettings["ConvCost"]);
                List<FP> linguaggio = new List<FP>();
                FP tmp = new FP();
                int i = 0;
                foreach (string item in pass)
                {
                    if (i == 4)
                    {
                        i = 0;
                    }

                    if (i == 1)
                    {
                        tmp.Linguaggio = item;
                    }

                    if (i == 3)
                    {
                        tmp.Dato = int.Parse(item);
                        totLine += tmp.Dato;
                    }

                    i++;
                    if (i == 4)
                    {
                        linguaggio.Add(tmp);
                        tmp = new FP();
                    }
                }

                double valoreFP = AnalisiFP(linguaggio, fp);
                TFS.FPTotali = valoreFP;
                TFS.CostiTotali = totLine;

            }
        }

        /// <summary>
        /// controllo readOnly
        /// </summary>
        /// <param name="f">
        /// files attributes
        /// </param>
        /// <param name="path">
        /// path da controllare e settare
        /// </param>
        public static void DecipherAttributes(FileAttributes f, string path)
        {
            // To set use File.SetAttributes
            File.SetAttributes(path, FileAttributes.ReadOnly);

            if ((f & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                Console.WriteLine("ReadOnly");
            }

            // To remove readonly use "-="
            f &= ~FileAttributes.ReadOnly;

            if ((f & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                Console.WriteLine("ReadOnly");
            }
            else
            {
                Console.WriteLine("Not ReadOnly");
            }
        }

        /// <summary>
        /// cancellazione file xml
        /// </summary>
        /// <param name="pippo">
        /// path del file da cancellare
        /// </param>
        public static void DelXml(string pippo)
        {
            if (File.Exists(pippo))
            {
                FileInfo di = new FileInfo(pippo);
                di.Delete();
            }
        }

        /// <summary>
        /// lista TP da un xml
        /// </summary>
        /// <returns>
        /// lista TP dall'xml
        /// </returns>
        public static List<TP> ListTp()
        {
            List<TP> listaTp = new List<TP>();
            if (File.Exists(Directory.GetCurrentDirectory() + "/lista.xml"))
            {
                XmlTextReader reader = new XmlTextReader(Directory.GetCurrentDirectory() + "/lista.xml");
                while (reader.Read())
                {
                    if (reader.Name.Equals("progetto") && reader.NodeType.Equals(XmlNodeType.Element))
                    {
                        listaTp.Add(new TP(reader.GetAttribute("name"), reader.GetAttribute("main"), reader.GetAttribute("CISoftware")));
                    }
                }
            }

            return listaTp;
        }

        /// <summary>
        /// metodo di pulizia della directory dai file autogenerati o non interessanti per il calcolo
        /// </summary>
        /// <param name="dir">
        /// directory di interesse
        /// </param>
        public static void PuliziaDirectory(string dir)
        {
            string sLine = string.Empty;
            DirectoryInfo di = Directory.CreateDirectory(dir);
            if (File.Exists("noDir.txt"))
            {
                StreamReader sr1 = new StreamReader("noDir.txt");
                while (sLine != null)
                {
                    sLine = sr1.ReadLine();
                    if (sLine != null)
                    {
                        DirectoryInfo[] controllo = di.GetDirectories(sLine, SearchOption.AllDirectories);
                        foreach (DirectoryInfo finfo in controllo)
                        {
                            var elenco = finfo.EnumerateFiles("*", SearchOption.AllDirectories);
                            foreach (var file in elenco)
                            {
                                file.IsReadOnly = false;
                            }

                            Directory.Delete(finfo.FullName, true);
                        }
                    }
                }
            }
            sLine = string.Empty;
            if (File.Exists("noFile.txt"))
            {
                StreamReader sr = new StreamReader("noFile.txt");
                while (sLine != null)
                {
                    sLine = sr.ReadLine();
                    if (sLine != null)
                    {
                        FileInfo[] cont = di.GetFiles(sLine, SearchOption.AllDirectories);
                        foreach (FileInfo f in cont)
                        {
                            f.IsReadOnly = false;
                            f.Delete();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// recupero dati per i FP
        /// </summary>
        /// <returns>Lista dei dati relativi agli FP</returns>
        public static List<FP> RecuperaFP()
        {
            List<FP> fp = new List<FP>();
            XmlTextReader reader = new XmlTextReader("fp_Data.xml");
            while (reader.Read())
            {
                if (reader.Name.Equals("linguaggio") && reader.NodeType.Equals(XmlNodeType.Element))
                {
                    FP tmp = new FP();
                    while (reader.MoveToNextAttribute())
                    {
                        if (reader.Name.Equals("name"))
                        {
                            tmp.Linguaggio = reader.Value;
                        }
                        else if (reader.Name.Equals("data"))
                        {
                            tmp.Dato = int.Parse(reader.Value);
                        }
                    }

                    fp.Add(tmp);
                }
            }

            reader.Close();
            return fp;
        }

        /// <summary>
        /// metodo di trasformazione da xml in excel
        /// </summary>
        public static void TrasformaInWI(String tfsName)
        {
            string repDir = System.Configuration.ConfigurationManager.AppSettings["repDir"];
            try
            {
                string proj = string.Empty;
                string coll = string.Empty;
                string cis = string.Empty;
                List<string> pass = new List<string>();
                DirectoryInfo di = Directory.CreateDirectory(repDir);
                if (di.EnumerateFiles().Count() == 0)
                {
                    using (TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(tfsName + coll), TFS.Cred))
                    {
                        TFS.CreateWI(tfs, pass, proj, proj);
                    }
                }
                else
                {
                    foreach (FileInfo finfo in di.EnumerateFiles())
                    {
                        if (finfo.Extension.Equals(".xml"))
                        {
                            using (XmlTextReader reader = new XmlTextReader(finfo.FullName))
                            {
                                while (reader.Read())
                                {
                                    if (reader.Name.Equals("TeamProject") && reader.NodeType.Equals(XmlNodeType.Element))
                                    {
                                        reader.Read();
                                        proj = reader.Value;
                                    }

                                    if (reader.Name.Equals("Collection") && reader.NodeType.Equals(XmlNodeType.Element))
                                    {
                                        reader.Read();
                                        coll = reader.Value;
                                    }

                                    if (reader.Name.Equals("CISoftware") && reader.NodeType.Equals(XmlNodeType.Element))
                                    {
                                        reader.Read();
                                        cis = reader.Value;
                                    }

                                    if (reader.Name.Equals("language") && reader.NodeType.Equals(XmlNodeType.Element))
                                    {
                                        int inizio = finfo.Name.IndexOf("#") + 1;
                                        int fine = finfo.Name.Length - 4;
                                        pass.Add(finfo.Name.Substring(inizio, fine - inizio));
                                        while (reader.MoveToNextAttribute())
                                        {
                                            if (reader.Name.Equals("name") || reader.Name.Equals("files_count") || reader.Name.Equals("code"))
                                            {
                                                pass.Add(reader.Value);
                                            }
                                        }

                                        pass.Add(TFS.Utenti(tfsName,proj, coll).ToString());
                                    }

                                    // if(!coll.Equals("") && !proj.Equals(""))
                                }
                            }
                        }

                        if (proj != string.Empty && pass.Count != 0)
                        {
                            using (TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(tfsName + coll), TFS.Cred))
                            {
                                TFS.CreateWI(tfs, pass, proj, cis);
                            }
                        }

                        proj = string.Empty;
                        coll = string.Empty;
                        pass = new List<string>();
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("errore in scrittura GlobalList: {0}", e.Message)));
                Console.WriteLine(e.Message);
                throw e;
            }
            finally
            {
                DirectoryInfo di = Directory.CreateDirectory(repDir);
                foreach (FileInfo finfo in di.EnumerateFiles())
                {
                    if (finfo.Extension.Equals(".xml"))
                    {
                        File.Delete(finfo.FullName);
                    }
                }
            }
        }

        /// <summary>
        /// Metodo di trasformazione da xml in excel
        /// </summary>
        /// <param name="excel">
        /// excel da riempire..
        /// </param>
        public static void TrasformaXML(String tfsName,Excel2 excel)
        {
            string repDir = System.Configuration.ConfigurationManager.AppSettings["repDir"];
            string proj = string.Empty;
            string coll = string.Empty;
            List<string> pass = new List<string>();
            DirectoryInfo di = Directory.CreateDirectory(repDir);
            foreach (FileInfo finfo in di.EnumerateFiles())
            {
                if (finfo.Extension.Equals(".xml"))
                {
                    XmlTextReader reader = new XmlTextReader(finfo.FullName);
                    while (reader.Read())
                    {
                        if (reader.Name.Equals("TeamProject") && reader.NodeType.Equals(XmlNodeType.Element))
                        {
                            reader.Read();
                            proj = reader.Value;
                        }

                        if (reader.Name.Equals("Collection") && reader.NodeType.Equals(XmlNodeType.Element))
                        {
                            reader.Read();
                            coll = reader.Value;
                        }

                        if (reader.Name.Equals("language") && reader.NodeType.Equals(XmlNodeType.Element))
                        {
                            pass = new List<string>();
                            int inizio = finfo.Name.IndexOf("#") + 1;
                            int fine = finfo.Name.Length - 4;
                            pass.Add(finfo.Name.Substring(inizio, fine - inizio));
                            while (reader.MoveToNextAttribute())
                            {
                                if (reader.Name.Equals("name") || reader.Name.Equals("files_count") || reader.Name.Equals("code"))
                                {
                                    pass.Add(reader.Value);
                                }
                            }

                            pass.Add(TFS.Utenti(tfsName, proj, coll).ToString());
                            pass.Insert(0, coll + "/" + proj);
                            excel.inserisciRiga(pass);
                        }

                        // if(!coll.Equals("") && !proj.Equals(""))
                    }
                }
            }
        }
    }
}