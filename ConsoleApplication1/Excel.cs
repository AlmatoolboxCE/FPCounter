//----------------------------------------------------------------------------------------------
// <copyright file="Excel.cs" company="Almaviva TSF" author="Gabriele Mazzocca">
// Copyright (c) Almaviva TSF.  All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------
namespace ConsoleAppTFS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    using CarlosAg.ExcelXmlWriter;

    /// <summary>
    /// Classe che gestisce i report su Excel.
    /// </summary>
    public class Excel
    {
        /// <summary>
        /// Cartella del file.
        /// </summary>
        public static string tmp = "report";
        string repDir = System.Configuration.ConfigurationManager.AppSettings["repDir"];
        string tmpDir = System.Configuration.ConfigurationManager.AppSettings["tmpDir"];

        /// <summary>
        /// Il workbook del file.
        /// </summary>
        Workbook book = new Workbook();

        /// <summary>
        /// Lista delle intestazioni delle colonne.
        /// </summary>
        List<string> col = new List<string>();

        /// <summary>
        /// True se il file è creato.
        /// </summary>
        bool creato = false;

        /// <summary>
        /// Nome del file.
        /// </summary>
        string filename = string.Empty;

        /// <summary>
        /// Sheet del worksheet.
        /// </summary>
        private Worksheet sheet = null;

        /// <summary>
        /// Costruttore di default.
        /// </summary>       
        public Excel()
        {
            this.col.Add("TeamProject");
            this.col.Add("controllo brench main");
            this.col.Add("Last Checkin");
            this.col.Add("controllo brench dev");
            this.col.Add("Last Checkin");
            this.col.Add("controllo brench rel");
            this.col.Add("Last Checkin");
            this.col.Add("Nr Task");
            this.col.Add("Nr Bug");
            this.col.Add("Nr All WI");
            this.col.Add("Nr Changesets");
        }

        /// <summary>
        /// Cartella del file.
        /// </summary>
        public static string Tmp
        {
            get { return Excel.tmp; }
            set { Excel.tmp = value; }
        }

        /// <summary>
        /// Aggiunge intestazioni di colonne al foglio.
        /// </summary>
        /// <param name="ins">
        /// Intestazioni da Inserire.
        /// </param>
        public void InsertCell(List<string> ins)
        {
            if (!this.creato)
            {
                this.col.Add("Applicazione");
                this.col.Add("Data");
                this.col.Add("Linguaggio");
                this.col.Add("Files");
                this.col.Add("Righe");
                this.col.Add("Utenti");
                this.SetSheet("report");
                this.creato = true;
            }

            WorksheetRow row = this.sheet.Table.Rows.Add();
            foreach (string data in ins)
            {
                row.Cells.Add(data);
            }
        }

        /// <summary>
        /// Inserisce nel foglio excel una nuova riga con i risultati parziali della scansione.
        /// </summary>
        /// <param name="collectionName">
        /// Il nome della collezione.
        /// </param>
        /// <param name="teamProjectName">
        /// Il nome del team project.
        /// </param>
        /// <param name="mainLinePath">
        /// Il path della mainline.
        /// </param>
        /// <param name="devPath">
        /// Il path del branch di DEV.
        /// </param>
        /// <param name="relPath">
        /// Il path del branch di REL.
        /// </param>
        /// <param name="numOfTasks">
        /// Il numero di Task.
        /// </param>
        /// <param name="numOfBugs">
        /// Il numero di Bug.
        /// </param>
        /// <param name="numOfWorkItems">
        /// Il numero dei work Items totale.
        /// </param>
        /// <param name="numOfChangesets">
        /// Il numero dei changeset.
        /// </param>
        /// <param name="lastCechInMain">
        /// L'ultimo check-in su Main.
        /// </param>
        /// <param name="lastCechIDev">
        /// L'ultimo check-in su Dev.
        /// </param>
        /// <param name="lastCechInRel">
        /// L'ultimo check-in su Rel.
        /// </param>
        public void InsertRow(string collectionName, string teamProjectName, string mainLinePath, string devPath, string relPath, int numOfTasks, int numOfBugs, int numOfWorkItems, int numOfChangesets, string lastCechInMain, string lastCechIDev, string lastCechInRel)
        {
            if (!collectionName.Equals(tmp))
            {
                this.SetSheet(collectionName);
                tmp = collectionName;
            }

            WorksheetRow row = this.sheet.Table.Rows.Add();
            row.Cells.Add(teamProjectName);
            row.Cells.Add(mainLinePath);
            row.Cells.Add(lastCechInMain);
            row.Cells.Add(devPath);
            row.Cells.Add(lastCechIDev);
            row.Cells.Add(relPath);
            row.Cells.Add(lastCechInRel);
            row.Cells.Add(numOfTasks.ToString());
            row.Cells.Add(numOfBugs.ToString());
            row.Cells.Add(numOfWorkItems.ToString());
            row.Cells.Add(numOfChangesets.ToString());
        }

        /// <summary>
        /// Salva il file.
        /// </summary>
        public void Salva()
        {
            // Save the file and open it
            this.book.Save(this.filename);
        }

        /// <summary>
        /// Aggiunge un nuovo sheet al workSheet.
        /// </summary>
        /// <param name="sheetName">
        /// Il nome dello sheet da aggiungere.
        /// </param>
        public void SetSheet(string sheetName)
        {
            this.sheet = this.book.Worksheets.Add(sheetName);

            // we can optionally set some column settings
            this.sheet.Table.Columns.Add(new WorksheetColumn(150));
            this.sheet.Table.Columns.Add(new WorksheetColumn(150));
            this.sheet.Table.Columns.Add(new WorksheetColumn(150));
            this.sheet.Table.Columns.Add(new WorksheetColumn(150));
            this.sheet.Table.Columns.Add(new WorksheetColumn(150));
            this.sheet.Table.Columns.Add(new WorksheetColumn(150));
            this.sheet.Table.Columns.Add(new WorksheetColumn(150));
            this.sheet.Table.Columns.Add(new WorksheetColumn(50));
            this.sheet.Table.Columns.Add(new WorksheetColumn(50));
            this.sheet.Table.Columns.Add(new WorksheetColumn(50));
            this.sheet.Table.Columns.Add(new WorksheetColumn(80));

            // Add a Worksheet with some data
            WorksheetRow row = this.sheet.Table.Rows.Add();
            foreach (string t in this.col)
            {
                row.Cells.Add(new WorksheetCell(t, "HeaderStyle"));
            }
        }

        /// <summary>
        /// Crea il file excel.
        /// </summary>
        /// <param name="titolo">
        /// Il titolo dell'XLS.
        /// </param>
        public void XLS(string titolo)
        {
            string user = System.Configuration.ConfigurationManager.AppSettings["user"];
            this.filename = repDir + titolo + "_" + DateTime.Today.Month + "-" + DateTime.Today.Year + ".xls";

            // Specify which Sheet should be opened and the size of window by default
            this.book.ExcelWorkbook.ActiveSheetIndex = 1;
            this.book.ExcelWorkbook.WindowTopX = 100;
            this.book.ExcelWorkbook.WindowTopY = 200;
            this.book.ExcelWorkbook.WindowHeight = 7000;
            this.book.ExcelWorkbook.WindowWidth = 8000;

            // Some optional properties of the Document
            this.book.Properties.Author = user;
            this.book.Properties.Title = titolo;
            this.book.Properties.Created = DateTime.Now;

            // Add some styles to the Workbook
            WorksheetStyle style = this.book.Styles.Add("HeaderStyle");
            style.Font.FontName = "Tahoma";
            style.Font.Size = 10;
            style.Font.Bold = true;
            style.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            style.Font.Color = "Black";

            // style.Interior.Color = "Blue";
            //            style.Interior.Pattern = StyleInteriorPattern.DiagCross;

            // Create the Default Style to use for everyone
            style = this.book.Styles.Add("Default");
            style.Font.FontName = "Tahoma";
            style.Font.Size = 10;

            // Generate 30 rows
            // InsertRow("","");
            // salva();
            // Process.Start(filename);
        }
    }
}