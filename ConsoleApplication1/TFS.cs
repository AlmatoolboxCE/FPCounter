//----------------------------------------------------------------------------------------------
// <copyright file="TFS.cs" company="Almaviva TSF" author="Gabriele Mazzocca  g.mazzocca@almaviva.it">
// Copyright (c) Almaviva TSF.  All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------
namespace ConsoleAppTFS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Text;
    using System.Xml;

    using AlmaLogger;

    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.Framework.Client;
    using Microsoft.TeamFoundation.Framework.Common;
    using Microsoft.TeamFoundation.Proxy;
    using Microsoft.TeamFoundation.Server;
    using Microsoft.TeamFoundation.VersionControl.Client;
    using Microsoft.TeamFoundation.VersionControl.Common;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;

    /// <summary>
    /// classe di gestione TFS
    /// </summary>
    public class TFS
    {
        /// <summary>
        /// lista di changeset senza WI
        /// </summary>
        public static List<Change_Data> CsNoWI = new List<Change_Data>();

        /// <summary>
        /// credenziali accesso TFS
        /// </summary>
        private static NetworkCredential cred = new NetworkCredential(User, Passwd, Dominio);

        /// <summary>
        /// dominio da caricare da configurazione
        /// </summary>
        private static string dominio = System.Configuration.ConfigurationManager.AppSettings["dominio"];

        /// <summary>
        /// Lista Q_S
        /// </summary>
        private static string listaQ_S;

        /// <summary>
        /// psw di accesso da caricare da conf
        /// </summary>
        private static string passwd = System.Configuration.ConfigurationManager.AppSettings["psw"];

        /// <summary>
        /// livello di ricorsione(0)
        /// </summary>
        private static int recursionlevel = 0;

        /// <summary>
        /// lista di ritorno
        /// </summary>
        private static List<string> ritorno = new List<string>();

        /// <summary>
        /// path del server TFS
        /// </summary>
        //private static string tfsName = System.Configuration.ConfigurationManager.AppSettings["TFSServer"];

        /// <summary>
        /// credenziali di accesso a tfs
        /// </summary>
        private static string user = System.Configuration.ConfigurationManager.AppSettings["user"];

        /// <summary>
        /// Gets or sets Costi totali(Linee di codice totali) da salvare su WI
        /// </summary>
        public static double CostiTotali
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets Credenziali accesso
        /// </summary>
        public static NetworkCredential Cred
        {
            get { return new NetworkCredential(User, Passwd, Dominio); }
            set { TFS.cred = new NetworkCredential(User, Passwd, Dominio); }
        }

        /// <summary>
        /// Gets or sets Dominio
        /// </summary>
        public static string Dominio
        {
            get { return TFS.dominio; }
            set { TFS.dominio = value; }
        }

        /// <summary>
        /// Gets or sets Function Point Totali
        /// </summary>
        public static double FPTotali
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets listaQ_S
        /// </summary>
        public static string ListaQ_S
        {
            get { return TFS.listaQ_S; }
            set { TFS.listaQ_S = value; }
        }

        /// <summary>
        /// Gets or sets PSW
        /// </summary>
        public static string Passwd
        {
            get { return TFS.passwd; }
            set { TFS.passwd = value; }
        }

        /// <summary>
        /// Gets or sets user
        /// </summary>
        public static string User
        {
            get { return TFS.user; }
            set { TFS.user = value; }
        }

        /// <summary>
        /// numero changeset per TP
        /// </summary>
        /// <param name="tfs">
        /// server TFS
        /// </param>
        /// <param name="tp">
        /// TP in oggetto
        /// </param>
        /// <returns>
        /// numero Changeset
        /// </returns>
        public static CSCount AllChangeset(TfsTeamProjectCollection tfs, string tp)
        {
            VersionControlServer versionControl;
            CsNoWI = new List<Change_Data>();

            // string TFSServerPath = @"http:// tfsfarm.tsf.local:8080/tfs/" + collection;
            // TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(TFSServerPath), Cred);
            tfs.Authenticate();
            versionControl = (VersionControlServer)tfs.GetService(typeof(VersionControlServer));

            // here you specify the exact time period you need to get the change set (from Date,to Date)
            VersionSpec fromDateVersion = new DateVersionSpec(DateTime.Now.AddMonths(-1));
            VersionSpec toDateVersion = new DateVersionSpec(DateTime.Now);

            // using versionControl API you can query histroy for changes set for (specific,all) User
            var data = versionControl.QueryHistory("$/" + tp, VersionSpec.Latest, 0, RecursionType.Full, null, fromDateVersion, toDateVersion, int.MaxValue, true, true);

            int i = 0;
            int old = 0;
            foreach (Microsoft.TeamFoundation.VersionControl.Client.Changeset pippo in data)
            {
                old++;
                if (pippo.WorkItems.Length == 0)
                {
                    CsNoWI.Add(new Change_Data(pippo.ChangesetId.ToString(), pippo.CommitterDisplayName, pippo.CreationDate.ToShortDateString(), pippo.Comment));
                    i++;
                }
            }

            CSCount csc = new CSCount();
            csc.All = old;
            csc.NonAssociato = i;
            return csc;
        }

        /// <summary>
        /// elenco competo WI
        /// </summary>
        /// <param name="tfs">
        /// server TFS
        /// </param>
        /// <param name="collection">
        /// collezione di ricerca
        /// </param>
        /// <param name="tp">
        /// Progetto di ricerca
        /// </param>
        /// <returns>
        /// numero wi per TP
        /// </returns>
        public static int AllWI(TfsTeamProjectCollection tfs, string collection, string tp)
        {
            List<string> lstQueries = new List<string>();

            // var tfs = new TfsTeamProjectCollection(new Uri("http:// tfsfarm.tsf.local:8080/tfs/" + collection), Cred);
            var wiStore = tfs.GetService<WorkItemStore>();
            WorkItemCollection wicollection = wiStore.Query("select [System.Id], [Microsoft.VSTS.Common.Priority], [System.State], [Microsoft.VSTS.Scheduling.RemainingWork], [Microsoft.VSTS.Scheduling.CompletedWork], [System.Title] from WorkItems where [System.TeamProject] = '" + tp + "'");

            // foreach (WorkItem select in wicollection)
            // {
            //    String titolo = select.Title;
            //    String tipo = select.Type.Name;
            // }
            return wicollection.Count;
        }

        /// <summary>
        /// controllo branches
        /// </summary>
        public static void BranchControl(String tfsName)
        {
            var tp = "AlmaToolBox";

            // 1.Construct the server object
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(tfsName + "/Almaviva"));
            VersionControlServer vcs = tfs.GetService<VersionControlServer>();

            // BranchObject[] bos=vcs.QueryBranchObjects(new ItemIdentifier("$/"+tp,VersionSpec.Latest),RecursionType.Full);
            // 2.Query all root branches
            BranchObject[] bos = vcs.QueryRootBranchObjects(RecursionType.OneLevel);

            // 3.Display all the root branches
            Array.ForEach(bos, (bo) => DisplayAllBranches(bo, vcs, tp));
            foreach (string p in ritorno)
            {
                if (p.Equals("\t"))
                {
                    Console.Write(p);
                }
                else
                {
                    Console.WriteLine(p);
                }
            }

            Console.ReadKey();
            Environment.Exit(0);
        }

        /// <summary>
        /// accesso alle solution c# del progetto
        /// </summary>
        /// <param name="server">
        /// server TFS
        /// </param>
        /// <param name="project">
        /// progetto in oggetto
        /// </param>
        /// <returns>
        /// lista di item relativi alle solution
        /// </returns>
        public static List<ExtendedItem> Button1_Click(TfsTeamProjectCollection server, string project)
        {
            try
            {
                // get project folder structure
                var path = "$/" + project;
                var sourceControl = (VersionControlServer)server.GetService(typeof(VersionControlServer));

                var pattern = "*.sln";
                var miItemSpec = new ItemSpec[] { new ItemSpec(path + "/" + pattern, RecursionType.Full) };

                // Get the latest version of the information for the items.
                var items = sourceControl.GetExtendedItems(miItemSpec, DeletedState.NonDeleted, ItemType.File)[0];

                return items.ToList();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                Logger.Fatal(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("errore in fase di cancellazione Workspace: {0}", exception.Message)));
                throw exception;
            }
        }

        /// <summary>
        /// controllo Global List
        /// </summary>
        /// <param name="tfs">
        /// connessione a TFS
        /// </param>
        /// <param name="project">
        /// progetto da analizzare
        /// </param>
        /// <param name="cis">
        /// ci software
        /// </param>
        /// <returns>
        /// true o false in base all'esito
        /// </returns>
        public static bool ControlloGlobalList(TfsTeamProjectCollection tfs, string project, string cis)
        {
            var store = new WorkItemStore(tfs);
            var globallists = store.ExportGlobalLists();
            XmlNode globalList = globallists.SelectSingleNode("//GLOBALLIST[@name='CIS']");

            XmlNodeList xml = globalList.ChildNodes;
            foreach (XmlNode x in xml)
            {
                if (x.Attributes[0].Value.Equals(cis))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// controllo padri sprint
        /// </summary>
        /// <param name="tfs">
        /// connnessione a TFS
        /// </param>
        /// <param name="tp">
        /// tp da analizzare
        /// </param>
        /// <returns>
        /// true verificato false altrimenti
        /// </returns>
        public static bool ControlloSprint(TfsTeamProjectCollection tfs, string tp)
        {
            string wi = "Sprint";
            List<string> lstQueries = new List<string>();
            var wiStore = tfs.GetService<WorkItemStore>();
            WorkItemCollection wicollection = wiStore.Query("select [System.Id], [Microsoft.VSTS.Common.Priority], [System.State], [Microsoft.VSTS.Scheduling.RemainingWork], [Microsoft.VSTS.Scheduling.CompletedWork], [System.Title] from WorkItems where [System.TeamProject] = '" + tp + "' and [System.WorkItemType] = '" + wi + "'");
            if (wicollection.Count == 0)
            {
                return false;
            }

            foreach (WorkItem select in wicollection)
            {
                var p = select.Fields["Start Date"].Value;
                var p1 = select.Fields["Finish Date"].Value;
                if (p != null)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// controllo padri Task
        /// </summary>
        /// <param name="tfs">
        /// connessione a TFS
        /// </param>
        /// <param name="tp">
        /// tp da analizzare
        /// </param>
        /// <returns>
        /// true se verifica positiva fale altrimenti
        /// </returns>
        public static bool ControlloTask(TfsTeamProjectCollection tfs, string tp)
        {
            string wi = "Task";
            List<string> lstQueries = new List<string>();
            var wiStore = tfs.GetService<WorkItemStore>();
            WorkItemCollection wicollection = wiStore.Query("select [System.Id], [Microsoft.VSTS.Common.Priority], [System.State], [Microsoft.VSTS.Scheduling.RemainingWork], [Microsoft.VSTS.Scheduling.CompletedWork], [System.Title] from WorkItems where [System.TeamProject] = '" + tp + "' and [System.WorkItemType] = '" + wi + "'");
            if (wicollection.Count == 0)
            {
                return false;
            }

            foreach (WorkItem select in wicollection)
            {
                if (select.State.Equals("Removed"))
                {
                }
                else
                {
                    List<string> trovato = new List<string>();
                    int t = select.RelatedLinkCount;
                    WorkItemLinkCollection child = select.WorkItemLinks;
                    if (child.Count == 0)
                    {
                        return false;
                    }

                    foreach (WorkItemLink wil in child)
                    {
                        trovato.Add(wiStore.GetWorkItem(wil.TargetId).Type.Name);
                    }

                    if (!trovato.Contains("Bug") && !trovato.Contains("Product Backlog Item"))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// creazione WI
        /// </summary>
        /// <param name="tfs">
        /// connessione tfs in oggetto
        /// </param>
        /// <param name="items">
        /// lista di items, dati linguaggio
        /// </param>
        /// <param name="tp">
        /// progetto in oggetto
        /// </param>
        /// <param name="cIS">
        /// Configuration Item Software
        /// </param>
        public static void CreateWI(TfsTeamProjectCollection tfs, List<string> items, string tp, string cIS)
        {
            List<FP> fp = XML.RecuperaFP();
            DateTime now = DateTime.Now;
            double totAlma = 0;
            double costi = double.Parse(System.Configuration.ConfigurationManager.AppSettings["ConvCost"]);
            List<FP> linguaggio = new List<FP>();
            WorkItemStore store = tfs.GetService<WorkItemStore>();
            Project proj = store.Projects[tp];
            if (!proj.WorkItemTypes.Contains("QAReport"))
            {
                string xmlPath = "wit/QAReport.xml";
                System.IO.StreamReader xmlStreamReader =
                    new System.IO.StreamReader(xmlPath);
                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                xmlDoc.Load(xmlStreamReader);
                System.Xml.XmlElement firstElement = (System.Xml.XmlElement)xmlDoc.DocumentElement;
                proj.WorkItemTypes.Import(firstElement);
            }
            store.RefreshCache(true);
            proj = store.Projects[tp];
            WorkItemType type = proj.WorkItemTypes["QAReport"];
            WorkItem workItem = new WorkItem(type);
            workItem["Title"] = cIS + " " + now;
            //workItem["CISoftware"] = cIS;
            workItem["DataD"] = now;
            CSCount csc = TFS.AllChangeset(tfs, tp);
            workItem["NChangeset"] = csc.NonAssociato;
            workItem["AllChangeset"] = csc.All;
            int i = 0;
            if (items.Count == 0)
            {
                workItem["DescrizioneHtml"] = "<table width='100%'><tr><th style='border-bottom:1px solid black'>Data</th><th style='border-bottom:1px solid black'>Linguaggio</th><th style='border-bottom:1px solid black'>Nr Files</th><th style='border-bottom:1px solid black'>Righe di Codice</th><th style='border-bottom:1px solid black'>Utenti</th></tr><tr>";
            }
            else
            {
                workItem["DescrizioneHtml"] = "<table width='100%'><tr><th style='border-bottom:1px solid black'>Data</th><th style='border-bottom:1px solid black'>Linguaggio</th><th style='border-bottom:1px solid black'>Nr Files</th><th style='border-bottom:1px solid black'>Righe di Codice</th><th style='border-bottom:1px solid black'>Utenti</th></tr><tr>";
                FP tmp = new FP();
                foreach (string item in items)
                {
                    if (i == 5)
                    {
                        i = 0;
                        workItem["DescrizioneHtml"] += "</tr><tr>";
                    }

                    if (i == 1)
                    {
                        tmp.Linguaggio = item;
                    }

                    if (i == 3)
                    {
                        tmp.Dato = int.Parse(item);
                        totAlma += tmp.Dato;
                    }

                    workItem["DescrizioneHtml"] += "<td style='border-bottom:1px solid black'>" + item + "</td>";

                    // workItem.Attachments.Add(new Attachment(""));
                    i++;
                    if (i == 5)
                    {
                        linguaggio.Add(tmp);
                        tmp = new FP();
                    }
                }

                workItem["DescrizioneHtml"] += "</tr></table>";
                double valoreFP = XML.AnalisiFP(linguaggio, fp);
                workItem["FP"] = valoreFP;
                workItem["ValoreCustom"] = totAlma;
                workItem["FP Totali"] = FPTotali;
                workItem["Valore Totale"] = CostiTotali;
                workItem["ControlloTask"] = TFS.ControlloTask(tfs, tp).ToString();
                workItem["ControlloSprint"] = TFS.ControlloSprint(tfs, tp).ToString();
                if (TFS.CsNoWI.Count != 0)
                {
                    workItem["DescrizioneChange"] = "<table width='100%'><tr><th style='border-bottom:1px solid black'>ChangesetID</th><th style='border-bottom:1px solid black'>User</th><th style='border-bottom:1px solid black'>Data Creazione</th><th style='border-bottom:1px solid black'>Commento</th></tr>";
                    foreach (Change_Data item in TFS.CsNoWI)
                    {
                        workItem["DescrizioneChange"] += "<tr>";
                        workItem["DescrizioneChange"] += "<td style='border-bottom:1px solid black'><A href='" + tfs.Uri.AbsolutePath + "web/UI/Pages/Scc/ViewChangeset.aspx?changeset=" + item.Id + "'>" + item.Id + "</A></td>";
                        workItem["DescrizioneChange"] += "<td style='border-bottom:1px solid black'>" + item.User + "</td>";
                        workItem["DescrizioneChange"] += "<td style='border-bottom:1px solid black'>" + item.Data + "</td>";
                        if (string.IsNullOrEmpty(item.Comment))
                        {
                            workItem["DescrizioneChange"] += "<td style='border-bottom:1px solid black'> -- </td>";
                        }
                        else
                        {
                            workItem["DescrizioneChange"] += "<td style='border-bottom:1px solid black'>" + item.Comment + "</td>";
                        }

                        workItem["DescrizioneChange"] += "</tr>";
                    }

                    workItem["DescrizioneChange"] += "</table>";
                }
            }

            Array result = workItem.Validate().ToArray();
            try
            {
                workItem.Save();
            }
            catch (Exception e)
            {
                Logger.Error(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("errore creazione WI: {0}", e.Message)));
                throw e;
            }

            FPTotali = 0;
            CostiTotali = 0;
        }

        /// <summary>
        /// recupero dati dal WI CI
        /// </summary>
        /// <param name="listTP">
        /// lista tp da aggiornare
        /// </param>
        /// <param name="tfs">
        /// connessione TFS
        /// </param>
        /// <param name="tp">
        /// tp in oggetto
        /// </param>
        public static void FindCI(List<TP> listTP, TfsTeamProjectCollection tfs, Project tp)
        {
            List<string> lstQueries = new List<string>();
            var wiStore = tfs.GetService<WorkItemStore>();
            WorkItemCollection wicollection = wiStore.Query("select [System.Id], [Microsoft.VSTS.Common.Priority], [System.State], [Microsoft.VSTS.Scheduling.RemainingWork], [Microsoft.VSTS.Scheduling.CompletedWork], [System.Title] from WorkItems where [System.TeamProject] = '" + tp.Name + "' and [System.WorkItemType] = 'CI'");
            foreach (WorkItem select in wicollection)
            {
                int id = select.Id;
                string title = (string)select["Title"];
                string pathConteggioFP = (string)select["Path conteggio FP"];
                if (title.ToUpper().Equals(tp.Name.ToUpper()))
                {
                    listTP.Add(new TP(tp.Name, pathConteggioFP, title));
                }
                else
                {
                    listTP.Add(new TP(tp.Name, pathConteggioFP, tp.Name + "." + title));
                }
            }
        }

        /// <summary>
        /// numero WI di un determinato tipo
        /// </summary>
        /// <param name="tfs">
        /// server TFS
        /// </param>
        /// <param name="collection">
        /// collezione in oggetto
        /// </param>
        /// <param name="tp">
        /// progetto in oggetto
        /// </param>
        /// <param name="wi">
        /// tipo WI
        /// </param>
        public static void FindWI(TfsTeamProjectCollection tfs, string collection, string tp, string wi)
        {
            List<string> lstQueries = new List<string>();
            var wiStore = tfs.GetService<WorkItemStore>();
            WorkItemCollection wicollection = wiStore.Query("select [System.Id], [Microsoft.VSTS.Common.Priority], [System.State], [Microsoft.VSTS.Scheduling.RemainingWork], [Microsoft.VSTS.Scheduling.CompletedWork], [System.Title] from WorkItems where [System.TeamProject] = '" + tp + "' and [System.WorkItemType] = '" + wi + "'");

            foreach (WorkItem select in wicollection)
            {
                string titolo = select.Title;
                Console.WriteLine(select.Type.Name + "\t" + select.Title);
                WorkItemLinkCollection child = select.WorkItemLinks;
                LinkCollection lc = select.Links;
                foreach (Link l in lc)
                {
                    Console.Write("\t");
                    string tipo = l.ToString();
                    RegisteredLinkType rlt = l.ArtifactLinkType;
                    Console.WriteLine("\t" + tipo + "\t" + rlt);
                    Console.ReadKey();
                }

                if (child.Count != 0)
                {
                    foreach (WorkItemLink l in child)
                    {
                        Console.Write("\t");
                        string tipo = l.BaseType.ToString();
                        string nome_relazione = l.LinkTypeEnd.Name;
                        RegisteredLinkType rlt = l.ArtifactLinkType;
                        WorkItem wit = wiStore.GetWorkItem(l.TargetId);
                        Console.WriteLine(wit.Type.Name + "\t" + wit.Title + "\t" + nome_relazione + "\t" + tipo + "\t" + rlt);
                        Console.ReadKey();
                    }
                }
                else
                {
                    Console.WriteLine("\t WI senza link");
                }
            }

            Environment.Exit(0);
        }

        /// <summary>
        /// restituisce l'elenco delle collection
        /// </summary>
        /// <returns>
        /// lista collection
        /// </returns>
        public static List<string> GetCollection(String tfsName)
        {
            List<string> listBox1 = new List<string>();
            using (var configurationServer = new TfsConfigurationServer(new Uri(tfsName), TFS.Cred))
            {
                var prova = configurationServer.GetService<ITeamProjectCollectionService>();
                foreach (var projectInfo in prova.GetCollections())
                {
                    listBox1.Add(projectInfo.Name);
                }

                return listBox1;
            }
        }

        /// <summary>
        /// da una collection l'elenco dei TP
        /// </summary>
        /// <param name="stringa">
        /// collection di ricerca
        /// </param>
        /// <returns>
        /// lista progetti
        /// </returns>
        public static List<TP> GetTP(String tfsName, string stringa)
        {
            List<TP> listTP = new List<TP>();
            using (var tfs = new TfsTeamProjectCollection(new Uri(tfsName + stringa), Cred))
            {
                tfs.Authenticate();
                var wiStore = tfs.GetService<WorkItemStore>();

                var projectCollection = wiStore.Projects;

                foreach (Project project in projectCollection)
                {
                    FindCI(listTP, tfs, project);
                }

                return listTP;
            }

            // return listTP;
        }

        /// <summary>
        /// recupero QA
        /// </summary>
        /// <param name="tfs">
        /// connessione TFS
        /// </param>
        /// <param name="tp">
        /// progetto da analizzare
        /// </param>
        public static void RecuperoQA(TfsTeamProjectCollection tfs, string tp)
        {
            List<int> cancellazione = new List<int>();
            string wi = "QAReport";
            int mesi = int.Parse(System.Configuration.ConfigurationManager.AppSettings["Cancellazione_in_Mesi"]);
            DateTime now = DateTime.Now.AddMonths(0 - mesi);
            List<string> lstQueries = new List<string>();
            var wiStore = tfs.GetService<WorkItemStore>();

            // wiStore.DestroyWorkItems(;
            WorkItemCollection wicollection = wiStore.Query("select [System.Id],[Microsoft.VSTS.Common.Priority], [System.State], [Microsoft.VSTS.Scheduling.RemainingWork], [Microsoft.VSTS.Scheduling.CompletedWork], [System.Title] from WorkItems where [System.TeamProject] = '" + tp + "' and [System.WorkItemType] = '" + wi + "'");
            foreach (WorkItem select in wicollection)
            {
                DateTime item = (DateTime)select.Fields["DataD"].Value;
                if (item.CompareTo(now) < 0)
                {
                    cancellazione.Add(select.Id);
                }
            }

            var error = wiStore.DestroyWorkItems(cancellazione);
            wiStore.RefreshCache();
            wiStore.SyncToCache();
        }

        /// <summary>
        /// metodo ricorsivo che gestisce l'import del QAReport sul progetto
        /// </summary>
        /// <param name="tfs">
        /// connessione a TFS
        /// </param>
        /// <param name="listTP">
        /// lista tp da aggiornare
        /// </param>
        /// <param name="area">
        /// nodo area
        /// </param>
        /// <param name="project">
        /// project ispezionato
        /// </param>
        public static void Recursive(TfsTeamProjectCollection tfs, List<TP> listTP, NodeCollection area, Project project)
        {
            foreach (Node item in area)
            {
                if (item.Name.ToUpper().Contains("Q_S"))
                {
                    var wit = project.WorkItemTypes;
                    if (!wit.Contains("QAReport"))
                    {
                        XmlDocument xdoc = new XmlDocument();
                        xdoc.Load("./wit/QAReport.xml");
                        XmlElement xml = xdoc.DocumentElement;
                        project.WorkItemTypes.Import(xml);
                    }

                    TP temp;
                    if (item.Name.StartsWith("Q_S"))
                    {
                        temp = new TP(project.Name, "$" + item.Name.Replace("Q_S", string.Empty).Replace("=", "/"), project.Name);
                    }
                    else
                    {
                        item.Name.Remove(item.Name.IndexOf("Q_S"));
                        temp = new TP(project.Name, "$" + item.Name.Remove(0, item.Name.IndexOf("Q_S")).Replace("Q_S", string.Empty).Replace("=", "/"), project.Name + "." + item.Name.Substring(0, item.Name.IndexOf("Q_S") - 1));
                    }

                    listTP.Add(temp);
                }

                if (item.HasChildNodes)
                {
                    Recursive(tfs, listTP, item.ChildNodes, project);
                }
            }
        }

        /// <summary>
        /// scrittura sulla global list.
        /// prende l'esistente e aggiunge il nuovo
        /// </summary>
        /// <param name="collection">
        /// collection a cui collegarsi
        /// </param>
        /// <param name="tp">
        /// tp per cui scrivere la global list
        /// </param>
        public static void ScriviGlobalList(String tfsName, string collection, List<TP> tp)
        {
            using (var tfs = new TfsTeamProjectCollection(new Uri(tfsName + collection), Cred))
            {
                try
                {
                    List<string> cisSoftware = new List<string>();

                    // Console.WriteLine("Avviato servizio su " + server.Uri.ToString());
                    // Console.WriteLine("loggato come " + System.Security.Principal.WindowsIdentity.GetCurrent().Name);
                    var store = new WorkItemStore(tfs);

                    // Console.WriteLine("ho creato lo store");
                    var globallists = store.ExportGlobalLists();

                    // Console.WriteLine("ho scaricato le global list");
                    RemoveAllListItem(globallists, "//GLOBALLIST[@name='CIS']");

                    // Console.WriteLine("ho rimosso i vecchi CIS");
                    var nuove = AddListItems(globallists, "//GLOBALLIST[@name='CIS']", "LISTITEM", "non disponibile");
                    foreach (var cis in tp)
                    {
                        if (!cisSoftware.Contains(cis.CISoftware))
                        {
                            cisSoftware.Add(cis.CISoftware);
                            nuove = AddListItems(globallists, "//GLOBALLIST[@name='CIS']", "LISTITEM", cis.CISoftware);
                        }
                    }

                    // Console.WriteLine("ho creato le nuove global list");

                    // var xmlFileName = System.IO.Directory.GetCurrentDirectory() + @"\tmp.xml";
                    // globallists.Save(xmlFileName);
                    store.ImportGlobalLists(globallists.InnerXml);
                    store.RefreshCache();

                    // Console.WriteLine("ho importato le nuove global list su tfs " + server.Uri.ToString());
                    // Console.WriteLine("operazione completata");
                }
                catch (Exception ex)
                {
                    Logger.Error(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("errore in scrittura GlobalList: {0}", ex.Message)));
                    throw ex;
                }
            }
        }

        /// <summary>
        /// numero utenti per il dato TP
        /// </summary>
        /// <param name="tp">
        /// progetto da controllare
        /// </param>
        /// <param name="collection">
        /// collection del progetto
        /// </param>
        /// <returns>
        /// numero utenti
        /// </returns>
        public static int Utenti(String tfsName, string tp, string collection)
        {
            int count = 0;
            using (var tfs = new TfsTeamProjectCollection(new Uri(tfsName + collection), Cred))
            {
                tfs.Authenticate();
                IGroupSecurityService gss = (IGroupSecurityService)tfs.GetService(typeof(IGroupSecurityService));
                string u = tfs.GetService<ICommonStructureService>().GetProjectFromName(tp).Uri;
                Identity[] id = gss.ListApplicationGroups(tfs.GetService<ICommonStructureService>().GetProjectFromName(tp).Uri);
                foreach (Identity identity in id)
                {
                    if (identity.AccountName != "Readers")
                    {
                        Identity sids = gss.ReadIdentity(SearchFactor.AccountName, "[" + tp + "]\\" + identity.AccountName, QueryMembership.Expanded);
                        count += sids.Members.Length;
                    }
                }

                return count;
            }
        }

        /// <summary>
        /// gestore branch
        /// </summary>
        public static void VersionControl(String tfsName)
        {
            using (var tfs = new TfsTeamProjectCollection(new Uri(tfsName), Cred))
            {
                var vc = tfs.GetService<VersionControlServer>();
                ItemSpec[] itemSpec = new ItemSpec[1];

                itemSpec[0] = new ItemSpec(
            @"$/Formazione ALM/",
                                          RecursionType.Full);
                BranchHistoryTreeItem[][] branchHistoryTree = vc.GetBranchHistory(new ItemSpec[] { new ItemSpec(@"$/Formazione ALM/", RecursionType.Full) }, LatestVersionSpec.Latest);
                foreach (BranchHistoryTreeItem[] tia in branchHistoryTree)
                {
                    foreach (BranchHistoryTreeItem ti in tia)
                    {
                        Console.WriteLine(ti.Relative.ToString());
                    }
                }
            }
        }

        /// <summary>
        /// numero WI di un determinato tipo
        /// </summary>
        /// <param name="tfs">
        /// server TFS
        /// </param>
        /// <param name="collection">
        /// collezione in oggetto
        /// </param>
        /// <param name="tp">
        /// progetto in oggetto
        /// </param>
        /// <param name="wi">
        /// tipo WI
        /// </param>
        /// <returns>
        /// numero WI del tp
        /// </returns>
        public static int WI(TfsTeamProjectCollection tfs, string collection, string tp, string wi)
        {
            List<string> lstQueries = new List<string>();
            var wiStore = tfs.GetService<WorkItemStore>();
            WorkItemCollection wicollection = wiStore.Query("select [System.Id], [Microsoft.VSTS.Common.Priority], [System.State], [Microsoft.VSTS.Scheduling.RemainingWork], [Microsoft.VSTS.Scheduling.CompletedWork], [System.Title] from WorkItems where [System.TeamProject] = '" + tp + "' and [System.WorkItemType] = '" + wi + "'");

            // foreach (WorkItem select in wicollection)
            // {
            //    String titolo = select.Title;
            //    String tipo = select.Type.Name;
            // }
            return wicollection.Count;
        }

        /// <summary>
        /// Aggiunge elementi alla lista
        /// </summary>
        /// <param name="theDoc">
        /// xmlDocument di riferimento
        /// </param>
        /// <param name="globalListXpath">
        /// path del nodo da modificare
        /// </param>
        /// <param name="nodeName">
        /// nome dei nodi da aggiungere
        /// </param>
        /// <param name="toAdd">
        /// lista dei valori da aggiungere
        /// </param>
        /// <returns>
        /// Il nodo modificato
        /// </returns>
        private static XmlNode AddListItems(XmlDocument theDoc, string globalListXpath, string nodeName, string toAdd)
        {
            var globalList = theDoc.SelectSingleNode(globalListXpath);
            if (globalList == null)
            {
                var lists = theDoc.FirstChild;
                var addingNode = "GLOBALLIST";
                globalList = theDoc.CreateNode(System.Xml.XmlNodeType.Element, addingNode, null);
                var name = theDoc.CreateAttribute("name");
                name.Value = "CIS";
                globalList.Attributes.Append(name);
                lists.AppendChild(globalList);
            }

            var newNode = theDoc.CreateNode(System.Xml.XmlNodeType.Element, nodeName, null);
            var newAttribute = theDoc.CreateAttribute("value");
            newAttribute.Value = toAdd;
            newNode.Attributes.Append(newAttribute);
            globalList.AppendChild(newNode);

            return globalList;
        }

        /// <summary>
        /// Display all brannches
        /// </summary>
        /// <param name="bo">
        /// oggetto branch
        /// </param>
        /// <param name="vcs">
        /// version control server
        /// </param>
        /// <param name="tp">
        /// Team project
        /// </param>
        private static void DisplayAllBranches(BranchObject bo, VersionControlServer vcs, string tp)
        {
            // 0.Prepare display indentation
            for (int tabcounter = 0; tabcounter < recursionlevel; tabcounter++)
            {
                ritorno.Add("\t");
            }

            // 1.Display the current branch
            ritorno.Add(string.Format("{0}", bo.Properties.RootItem.Item));

            // 2.Query all child branches (one level deep)
            BranchObject[] childBos = vcs.QueryBranchObjects(bo.Properties.RootItem, RecursionType.OneLevel);

            // 3.Display all children recursively
            recursionlevel++;
            foreach (BranchObject child in childBos)
            {
                if (child.Properties.RootItem.Item == bo.Properties.RootItem.Item)
                {
                    continue;
                }

                DisplayAllBranches(child, vcs, tp);
            }

            recursionlevel--;
        }

        /// <summary>
        /// restituisce i Tag contenuti nell'Area del progetto ispeizonato
        /// </summary>
        /// <param name="listTP">
        /// lista dei TP da riempire
        /// </param>
        /// <param name="tfs">
        /// connessione a TFS
        /// </param>
        /// <param name="project">
        /// tp da ispezionare
        /// </param>
        private static void GetTag(List<TP> listTP, TfsTeamProjectCollection tfs, Project project)
        {
            // Area Path
            ListaQ_S = string.Empty;

            Recursive(tfs, listTP, project.AreaRootNodes, project);
            if (ListaQ_S != string.Empty)
            {
                Console.Write("\n" + project.Name + ": " + ListaQ_S);
                Console.WriteLine();
                Console.WriteLine(project.Name);
            }

            string[] listArea = ListaQ_S.Split("Q_S".ToCharArray());
            foreach (string i in listArea)
            {
                Console.Write(i.Replace('=', '/'));
            }
        }

        /// <summary>
        /// rimuove tutti gli elementi tranne il primo
        /// </summary>
        /// <param name="theDoc">
        /// xmlDocument di riferimento
        /// </param>
        /// <param name="globalListXpath">
        /// Xpath del nodo da modificare
        /// </param>
        private static void RemoveAllListItem(XmlDocument theDoc, string globalListXpath)
        {
            try
            {
                var globalList = theDoc.SelectSingleNode(globalListXpath);
                for (int i = 0; i < globalList.ChildNodes.Count; i++)
                {
                    globalList.RemoveChild(globalList.ChildNodes[i--]);
                }
            }
            catch (Exception ex)
            {
                Logger.Error(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("errore in rimozione GlobalList: {0}", ex.Message)));

                // Console.WriteLine("non sono riuscito a rimuovere i vecchi elementi. Probabilmente la lista non esiste");
                throw ex;
            }
        }
    }
}