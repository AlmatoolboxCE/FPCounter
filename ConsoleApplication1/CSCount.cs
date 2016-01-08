//----------------------------------------------------------------------------------------------
// <copyright file="CSCount.cs" company="Almaviva TSF" author="Gabriele Mazzocca  g.mazzocca@almaviva.it">
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
    /// classe relativa al conteggio dei changeset
    /// </summary>
    public class CSCount
    {
        /// <summary>
        /// numero changeset
        /// </summary>
        private int all;

        /// <summary>
        /// numero changeset non associati
        /// </summary>
        private int nonAssociato;

        /// <summary>
        /// numero changeset
        /// </summary>
        public int All
        {
            get { return this.all; }
            set { this.all = value; }
        }

        /// <summary>
        /// numero changeset non associati
        /// </summary>
        public int NonAssociato
        {
            get { return this.nonAssociato; }
            set { this.nonAssociato = value; }
        }
    }
}