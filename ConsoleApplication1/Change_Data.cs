//----------------------------------------------------------------------------------------------
// <copyright file="Change_Data.cs" company="Almaviva TSF" author="Gabriele Mazzocca  g.mazzocca@almaviva.it">
// Copyright (c) Almaviva TSF.  All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------
namespace ConsoleAppTFS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// classe di gestione dei Changeset
    /// </summary>
    public class Change_Data
    {
        /// <summary>
        /// Commenti del change
        /// </summary>
        private string comment;

        /// <summary>
        /// Data del change
        /// </summary>
        private string data;

        /// <summary>
        /// id del change
        /// </summary>
        private string id;

        /// <summary>
        /// User del change
        /// </summary>
        private string user;

        /// <summary>
        /// comment del change
        /// </summary>
        public string Comment
        {
            get { return this.comment; }
            set { this.comment = value; }
        }

        /// <summary>
        /// data changeset
        /// </summary>
        public string Data
        {
            get { return this.data; }
            set { this.data = value; }
        }

        /// <summary>
        /// Id del changeset
        /// </summary>
        public string Id
        {
            get { return this.id; }
            set { this.id = value; }
        }

        /// <summary>
        /// User che ha effettuato i change
        /// </summary>
        public string User
        {
            get { return this.user; }
            set { this.user = value; }
        }
        /// <summary>
        /// Inizializza un Change coi valori passati
        /// </summary>
        /// <param name="i">ID del change</param>
        /// <param name="u">User del change</param>
        /// <param name="d">data committer</param>
        /// <param name="c">commento change</param>
        public Change_Data( string i, string u, string d, string c)
        {
            Id = i;
            User = u;
            Data = d;
            Comment = c;
        }
    }
}