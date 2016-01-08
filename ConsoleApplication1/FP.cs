//----------------------------------------------------------------------------------------------
// <copyright file="TP.cs" company="Almaviva TSF" author="Gabriele Mazzocca  g.mazzocca@almaviva.it">
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
    /// il tp da osservare
    /// </summary>
    public class FP
    {
        /// <summary>
        /// path mainline
        /// </summary>
        private int dato;

        /// <summary>
        /// Gets nome Team project
        /// </summary>
        public int Dato
        {
            get { return dato; }
            set { dato = value; }
        }

        /// <summary>
        /// nome team project
        /// </summary>
        private string linguaggio;

        /// <summary>
        /// Gets path della main line
        /// </summary>
        public string Linguaggio
        {
            get { return linguaggio; }
            set { linguaggio = value; }
        }

        /// <summary>
        /// Inizializes a new instance of the TP
        /// </summary>
        /// <param name="n">
        /// nome del TP
        /// </param>
        /// <param name="m">
        /// main line del tp
        /// </param>
        /// <returns>
        /// il tp
        /// </returns>
        public FP(int d, string l)
        {
            this.Dato = d;
            this.Linguaggio = l;
        }

        public FP()
        {
        }
    }
}