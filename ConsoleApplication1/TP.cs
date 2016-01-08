//----------------------------------------------------------------------------------------------
// <copyright file="TP.cs" company="Almaviva TSF" author="Gabriele Mazzocca  g.mazzocca@almaviva.it">
// Copyright (c) Almaviva TSF.  All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------
namespace ConsoleAppTFS
{
    using System;
    using System.Collections.Generic;
    using System.Data.Entity;
    using System.Linq;
    using System.Text;
    using System.Xml;


    /// <summary>
    /// Il TP da osservare.
    /// </summary>
    public class TP
    {
        /// <summary>
        /// Nome CISoftware.
        /// </summary>
        private string ciSoftware;

        /// <summary>
        /// Path mainline.
        /// </summary>
        private string main;

        /// <summary>
        /// Nome team project.
        /// </summary>
        private string name;

        /// <summary>
        /// Costruttore classe TP.
        /// </summary>
        /// <param name="n">
        /// Nome del TP.
        /// </param>
        /// <param name="m">
        /// Main line del tp.
        /// </param>
        /// <param name="cIS">
        /// configuratio item software
        /// </param>
        /// <returns>
        /// Il tp.
        /// </returns>
        public TP(string n, string m, string cIS)
        {
            this.Name = n;
            this.Main = m;
            this.CISoftware = cIS;
        }

        /// <summary>
        /// Gets or sets nome CISoftware.
        /// </summary>
        public string CISoftware
        {
            get { return this.ciSoftware; }
            set { this.ciSoftware = value; }
        }

        /// <summary>
        /// Gets path della main line.
        /// </summary>
        public string Main
        {
            get { return this.main; }
            set { this.main = value; }
        }

        /// <summary>
        /// Gets nome Team project.
        /// </summary>
        public string Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        ///// <summary>
        ///// recupera i TP dal db di acna
        ///// </summary>
        //public static void GetAllTPfromACNA()
        //{
        //    // create new instance of XmlDocument
        //    XmlDocument doc = new XmlDocument();
        //    var root = doc.CreateElement("progetti");

        //    using (var db = new ACNAEntities())
        //    {
        //        List<SISTEMA> sys = db.SISTEMA.ToList();
        //        foreach (var s in sys)
        //        {
        //            ALM a = db.ALM.Find(s.ID_SISTEMA);
        //            string tp = a.TP_ALM.Split('/')[1];
        //            string collection = a.TP_ALM.Split('/')[0];
        //            if (string.IsNullOrWhiteSpace(s.MAIN_ALM))
        //            {
        //                List<SOTTOSISTEMA> sottosys = db.SOTTOSISTEMA.Where(so => so.ID_SISTEMA == s.ID_SISTEMA).ToList();
        //                if (sottosys.Count != 0)
        //                {
        //                    foreach (var sotto in sottosys)
        //                    {
        //                        if (sotto.MAIN_ALM.ToUpper().StartsWith("$/"))
        //                        {
        //                            XmlElement nodeSotto = doc.CreateElement("progetto");
        //                            nodeSotto.SetAttribute("name", tp);
        //                            nodeSotto.SetAttribute("collection", collection);
        //                            nodeSotto.SetAttribute("main", sotto.MAIN_ALM);
        //                            nodeSotto.SetAttribute("CISoftware", tp + "." + sotto.NOME_SOTTOSISTEMA);
        //                            root.AppendChild(nodeSotto);
        //                        }
        //                        else
        //                        {
        //                            List<COMPONENT> componente = db.COMPONENT.Where(cp => cp.ID_SOTTOSISTEMA == sotto.ID_SOTTOSISTEMA).ToList();

        //                            if (componente.Count != 0)
        //                            {
        //                                foreach (var comp in componente)
        //                                {
        //                                    if (comp.MAIN_ALM.ToUpper().StartsWith("$/"))
        //                                    {
        //                                        XmlElement nodeSotto = doc.CreateElement("progetto");
        //                                        nodeSotto.SetAttribute("name", tp);
        //                                        nodeSotto.SetAttribute("collection", collection);
        //                                        nodeSotto.SetAttribute("main", comp.MAIN_ALM);
        //                                        nodeSotto.SetAttribute("CISoftware", tp + "." + sotto.NOME_SOTTOSISTEMA);
        //                                        root.AppendChild(nodeSotto);
        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            else if (s.MAIN_ALM.ToUpper().Equals("X"))
        //            {
        //                XmlElement node = doc.CreateElement("progetto");
        //                node.SetAttribute("name", tp);
        //                node.SetAttribute("collection", collection);
        //                node.SetAttribute("main", string.Empty);
        //                node.SetAttribute("CISoftware", tp);
        //                root.AppendChild(node);
        //            }
        //            else if (s.MAIN_ALM.ToUpper().StartsWith("$/"))
        //            {
        //                XmlElement node = doc.CreateElement("progetto");
        //                node.SetAttribute("name", tp);
        //                node.SetAttribute("collection", collection);
        //                node.SetAttribute("main", s.MAIN_ALM);
        //                node.SetAttribute("CISoftware", tp);
        //                root.AppendChild(node);
        //            }
        //        }

        //        doc.AppendChild(root);

        //        // save back
        //        doc.Save("lista.xml");
        //    }

        //    return;
        //}
    }
}