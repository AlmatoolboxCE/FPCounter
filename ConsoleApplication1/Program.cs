//----------------------------------------------------------------------------------------------
// <copyright file="Program.cs" company="Almaviva TSF" author="Gabriele Mazzocca  g.mazzocca@almaviva.it">
// Copyright (c) Almaviva TSF.  All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------
namespace ConsoleAppTFS
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Management;
    using System.Text;

    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.VersionControl.Client;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    using AlmaLogger;

    /// <summary>
    /// programma gestione qualita del codice si TFS
    /// </summary>
    public class Program
    {
        /// <summary>
        /// prima della pendenza
        /// </summary>
        /// <param name="sender">
        /// sender chiamante
        /// </param>
        /// <param name="e">
        /// evento event
        /// </param>
        internal static void OnBeforeCheckinPendingChange(object sender, ProcessingChangeEventArgs e)
        {
            Console.WriteLine("Checking in " + e.PendingChange.LocalItem);
        }

        /// <summary>
        /// metodo gestione on get
        /// </summary>
        /// <param name="sender">
        /// sendere sender
        /// </param>
        /// <param name="e">
        /// evento event
        /// </param>
        internal static void OnGetting(object sender, GettingEventArgs e)
        {
            Console.WriteLine("Getting: " + e.TargetLocalItem + ", status: " + e.Status);
        }

        /// <summary>
        /// evento new pwnding change
        /// </summary>
        /// <param name="sender">
        /// chiamante sender
        /// </param>
        /// <param name="e">
        /// evento event
        /// </param>
        internal static void OnNewPendingChange(object sender, PendingChangeEventArgs e)
        {
            Console.WriteLine("Pending " + PendingChange.GetLocalizedStringForChangeType(e.PendingChange.ChangeType) +
                              " on " + e.PendingChange.LocalItem);
        }

        /// <summary>
        /// Non Fatal Error occorrente
        /// </summary>
        /// <param name="sender">
        /// instanza chiamante
        /// </param>
        /// <param name="e">
        /// eccezione rilevata
        /// </param>
        internal static void OnNonFatalError(object sender, ExceptionEventArgs e)
        {
            if (e.Exception != null)
            {
                Console.Error.WriteLine("Non-fatal exception: " + e.Exception.Message);
            }
            else
            {
                Console.Error.WriteLine("Non-fatal failure: " + e.Failure.Message);
            }
        }

        /// <summary>
        /// metodo main del progetto
        /// </summary>
        /// <param name="args">
        /// paramentri da console, non utilizzati
        /// </param>
        private static void Main(string[] args)
        {
            try
            {
                if (!Logger.Initialized)
                {
                    Logger.Init(Directory.GetCurrentDirectory() + "\\loginfo.config");
                }
            }
            catch (Exception e)
            {
                Logger.Fatal(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("errore in fase di inizializzazione: {0}", e.Message)));
            }

            string metriche = System.Configuration.ConfigurationManager.AppSettings["metriche"];
            string[] collections = System.Configuration.ConfigurationManager.AppSettings["collection"].Split(',');
            string repDir = System.Configuration.ConfigurationManager.AppSettings["repDir"];
            string tmpDir = System.Configuration.ConfigurationManager.AppSettings["tmpDir"];
            //            string tfsName = System.Configuration.ConfigurationManager.AppSettings["TFSServer"];
            string[] tfsServers = System.Configuration.ConfigurationManager.AppSettings["TFSServer"].Split(',');

            //TFS.BranchControl();
            // TFS.VersionControl();
            //            TP.GetAllTPfromACNA();
            //            List<TP> tp = TFS.GetTP();
            Workspace workspace = null;

            // Get a reference to our Team Foundation Server.
            string pathSI = string.Empty;

            // bool main = false;
            // bool dev = false;
            // bool rel = false;
            //string lcMain = string.Empty;
            //string lcDev = string.Empty;
            //string lcRel = string.Empty;
            //var allWI = 0;

            // modifica
            //var wiCountTask = 0;
            //var wiCountBug = 0;
            try
            {
                foreach (string tfsName in tfsServers)
                {

                    List<string> collection = TFS.GetCollection(tfsName);

                    foreach (string nome_coll in collection)
                    {
                        Excel.tmp = "report";
                        if (collections.Contains(nome_coll))
                        {
                            List<TP> tp = TFS.GetTP(tfsName,nome_coll);
                            //TFS.ScriviGlobalList(nome_coll, tp);
                            foreach (TP coll in tp)
                            {
                                using (TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(tfsName + "/" + nome_coll), TFS.Cred))
                                {
                                    TFS.RecuperoQA(tfs, coll.Name);
                                    TFS.ControlloSprint(tfs, coll.Name);
                                    // TFS.VersionControl();
                                    // TFS.CreateWI(tfs, coll);
                                    // allWI = TFS.AllWI(tfs, nome_coll, coll.Name);
                                    /*
                                     *Ricerca del numero di task in un tp 
                                     */

                                    //wiCountTask = TFS.WI(tfs, nome_coll, coll.Name, "Task");
                                    //wiCountBug = TFS.WI(tfs, nome_coll, coll.Name, "Bug");
                                    /*
                                     * Fine ricerca
                                     */

                                    // Get a reference to Version Control.
                                    VersionControlServer versionControl = (VersionControlServer)tfs.GetService(typeof(VersionControlServer));

                                    //// Listen for the Source Control events.
                                    versionControl.NonFatalError += Program.OnNonFatalError;
                                    versionControl.Getting += Program.OnGetting;
                                    versionControl.BeforeCheckinPendingChange += Program.OnBeforeCheckinPendingChange;
                                    versionControl.NewPendingChange += Program.OnNewPendingChange;

                                    try
                                    {
                                        System.Security.SecureString psw = new System.Security.SecureString();
                                        foreach (char c in TFS.Passwd)
                                        {
                                            psw.AppendChar(c);
                                        }

                                        if (metriche.Equals("si"))
                                        {
                                            // Creazione workspace.
                                            string tmp = coll.Name.Replace(" ", "_");
                                            workspace = versionControl.CreateWorkspace("QA_USER", versionControl.AuthorizedUser);
                                            if (workspace == null)
                                            {
                                                workspace = versionControl.CreateWorkspace("QA_USER", versionControl.AuthorizedUser);
                                            }

                                            // Create a mapping using the Team Project supplied on the command line.
                                            workspace.Map("$/" + coll.Name, tmpDir + tmp);
                                            WorkingFolder wf = workspace.GetWorkingFolderForServerItem("$/" + coll.Name);
                                            pathSI = wf.LocalItem;

                                            // Get the files from the repository.
                                            if (string.IsNullOrWhiteSpace(coll.Main))
                                            {
                                                VersionSpec fromDateVersion = new DateVersionSpec(DateTime.Now.AddMonths(-2));
                                                // workspace.Get(arrPath.ToArray(), VersionSpec.Latest, RecursionType.Full, GetOptions.GetAll);
                                                workspace.Get(fromDateVersion, GetOptions.GetAll);
                                                DirectoryInfo di1 = Directory.CreateDirectory(pathSI);
                                                DirectoryInfo[] directories = di1.GetDirectories("main*", SearchOption.AllDirectories);
                                                foreach (var dir in directories)
                                                {
                                                    string fullName = "\"" + dir.FullName.Replace('\\', '/') + "\"";
                                                    var outputXml = string.Empty;
                                                    if (coll.CISoftware != string.Empty)
                                                    {
                                                        outputXml = repDir + coll.CISoftware.Replace(" ", "_") + "-" + "#" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + ".xml";
                                                    }
                                                    else
                                                    {
                                                        outputXml = repDir + tmp + "-" + "#" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + ".xml";
                                                    }

                                                    /*
                                                     * controllo pre-pulizia directory
                                                     */
                                                    Logger.Info(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("Comando lanciato per il cloc: {0}", Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe" + fullName + @" --xml --out=" + outputXml)));
                                                    Console.WriteLine(Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe");
                                                    Console.WriteLine(fullName + @" --xml --out=" + outputXml);
                                                    Process p1 = Process.Start(Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe", fullName + @" --xml --out=" + outputXml, TFS.User, psw, TFS.Dominio);
                                                    while (!p1.HasExited)
                                                    {
                                                    }

                                                    XML.AppendTP(coll.Name, nome_coll, outputXml, coll.CISoftware);
                                                    XML.ControlloTotali();
                                                    XML.DelXml(outputXml);
                                                    /*
                                                     * controllo post-pulizia directory
                                                     */
                                                    XML.PuliziaDirectory(pathSI);
                                                    Process p = Process.Start(Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe", fullName + @" --xml --out=" + outputXml, TFS.User, psw, TFS.Dominio);
                                                    while (!p.HasExited)
                                                    {
                                                    }

                                                    XML.AppendTP(coll.Name, nome_coll, outputXml, coll.CISoftware);
                                                    XML.TrasformaInWI(tfsName);
                                                }
                                            }//nel caso in cui la main line è definita
                                            else
                                            {
                                                string[] arrPath = coll.Main.Split(',');

                                                //VersionSpec fromDateVersion = new DateVersionSpec(DateTime.Now.AddMonths(-2));
                                                // for (int i = 0; i < arrPath.Length; i++)
                                                // {
                                                //    workspace.Get(new GetRequest(arrPath[i], RecursionType.Full, VersionSpec.Latest), GetOptions.Remap);
                                                // }
                                                var errore = workspace.Get(arrPath, VersionSpec.Latest, RecursionType.Full, GetOptions.Remap);
                                                string fullName = "\"" + pathSI + "\"";
                                                var outputXml = string.Empty;
                                                if (coll.CISoftware != string.Empty)
                                                {
                                                    outputXml = repDir + coll.CISoftware.Replace(" ", "_") + "-" + "#" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + ".xml";
                                                }
                                                else
                                                {
                                                    outputXml = repDir + tmp + "-" + "#" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + ".xml";
                                                }

                                                /*
                                                 * controllo pre-pulizia directory
                                                 */
                                                Logger.Info(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("Comando lanciato per il cloc: {0}", Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe" + fullName + @" --xml --out=" + outputXml)));
                                                Console.WriteLine(Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe");
                                                Console.WriteLine(fullName + @" --xml --out=" + outputXml);
                                                Process p1 = Process.Start(Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe", fullName + @" --xml --out=" + outputXml, TFS.User, psw, TFS.Dominio);
                                                while (!p1.HasExited)
                                                {
                                                }

                                                XML.AppendTP(coll.Name, nome_coll, outputXml, coll.CISoftware);
                                                XML.ControlloTotali();
                                                XML.DelXml(outputXml);
                                                /*
                                                 * controllo post-pulizia directory
                                                 */
                                                XML.PuliziaDirectory(pathSI);
                                                Logger.Info(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("Comando lanciato per il cloc: {0}", Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe" + fullName + @" --xml --out=" + outputXml)));
                                                Process p = Process.Start(Directory.GetCurrentDirectory() + "/eseguibili/cloc.exe", fullName + @" --xml --out=" + outputXml, TFS.User, psw, TFS.Dominio);
                                                while (!p.HasExited)
                                                {
                                                }
                                                if (File.Exists(outputXml))
                                                {
                                                    XML.AppendTP(coll.Name, nome_coll, outputXml, coll.CISoftware);
                                                    XML.TrasformaInWI(tfsName);
                                                }
                                                else
                                                {
                                                    TFS.CreateWI(tfs, new List<string>(), coll.Name, coll.CISoftware);
                                                }

                                            }
                                        }

                                        // List<ExtendedItem> di = TFS.Button1_Click(tfs, coll.Name);
                                        // foreach (ExtendedItem dinfo in di)
                                        // {
                                        //    string prova = dinfo.SourceServerItem;
                                        //    prova = prova.ToLower();
                                        //    if (prova.Contains("main"))
                                        //    {
                                        //        main = true;
                                        //        lcMain = dinfo.CheckinDate.ToString();
                                        //    }

                                        // if (prova.Contains("dev"))
                                        //    {
                                        //        dev = true;
                                        //        lcDev = dinfo.CheckinDate.ToString();
                                        //    }

                                        // if (prova.Contains("rel"))
                                        //    {
                                        //        rel = true;
                                        //        lcRel = dinfo.CheckinDate.ToString();
                                        //    }
                                        // }
                                    }
                                    catch (Exception e)
                                    {
                                        Logger.Error(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("Errore sul controllo del progetto: {0}", e.Message)));
                                        Console.WriteLine(e.Message);
                                    }
                                    finally
                                    {
                                        try
                                        {
                                            if (metriche.Equals("si"))
                                            {
                                                // Delete the workspace.
                                                workspace.Delete();
                                                string tmp = pathSI.Replace(" ", "#");
                                                string dirObject = "Win32_Directory.Name='" + pathSI + "'";
                                                using (ManagementObject managementObject = new ManagementObject(dirObject))
                                                {
                                                    managementObject.Get();
                                                    ManagementBaseObject outParams = managementObject.InvokeMethod("Delete", null, null);

                                                    // ReturnValue should be 0, else failure
                                                    if (Convert.ToInt32(outParams.Properties["ReturnValue"].Value) != 0)
                                                    {
                                                        Console.Write("Errore");
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            Console.WriteLine(e.Message);
                                            Logger.Fatal(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("errore in fase di cancellazione Workspace: {0}", e.Message)));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Fatal(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "INT", string.Format("errore in fase di cancellazione Workspace: {0}", e.StackTrace)));

            }
        }
    }
}