using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelToSqlServer.Models
{
    public class GL
    {
        [Column("Année Mo")]
        public string AnneeMois { get; set; }

        [Column("Client")]
        public string Compte { get; set; }

        [Column("Texte")]
        public string Texte { get; set; }

        [Column("N° pièce")]
        public string NumPiece { get; set; }

        [Column("Date pièce")]
        public string DatePiece { get; set; }

        [Column("Affectation")]
        public string Affectation { get; set; }

        [Column("Référence")]
        public string Reference { get; set; }

        [Column("Typ")]
        public string TypeDePiece { get; set; }

        [Column("Mtant")]
        public Single MontantEnDeviseInterne { get; set; }

        [Column("EC")]
        public string PieceRapprochement { get; set; }
    }
}