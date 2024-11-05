using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendNewsLetter
{
    /// <summary>Holds concert data used by this application for setting of controls.
    /// <para>Data is retrieved from the season XML object corresponding to the file JazzProgramm_20xx_20yy.xml for the current season</para>
    /// <para></para>
    /// </summary>
    public class JazzConcert
    {
        /// <summary>Band name</summary>
        private string m_band_name = @"";
        /// <summary>Get and set band name</summary>
        public string BandName { get { return m_band_name; } set { m_band_name = value; } }

        /// <summary>Concert year as int</summary>
        private int m_concert_year = -12345;
        /// <summary>Get and set concert year as int</summary>
        public int Year{ get { return m_concert_year; } set { m_concert_year = value; } }

        /// <summary>Concert month as int</summary>
        private int m_concert_month = -12345;
        /// <summary>Get and set concert month as int</summary>
        public int Month { get { return m_concert_month; } set { m_concert_month = value; } }

        /// <summary>Concert day as int</summary>
        private int m_concert_day = -12345;
        /// <summary>Get and set concert day as int</summary>
        public int Day { get { return m_concert_day; } set { m_concert_day = value; } }

        /// <summary>Poster file name</summary>
        private string m_poster_file_name = @"";
        /// <summary>Get and set poster file name</summary>
        public string Poster { get { return m_poster_file_name; } set { m_poster_file_name = value; } }

        /// <summary>Premises name and address</summary>
        private string m_premises_address = @"";

        /// <summary>Get and set premises name and address</summary>
        public string Premises { get { return m_premises_address; } set { m_premises_address = value; } }

        /// <summary>Checks the member data</summary>
        public bool Check(out string o_error)
        {
            o_error = @"";

            if (BandName.Length == 0)
            {
                o_error = @"JazzConcert.Check BandName is not set";
            }

            if (Year < 1996)
            {
                o_error = @"JazzConcert.Check Year= " + Year.ToString() + @" < 1996";
            }

            if (Month < 1 || Month > 12)
            {
                o_error = @"JazzConcert.Check Month= " + Month.ToString() + @" < 1 or > 12";
            }

            if (Poster.Length == 0)
            {
                o_error = @"JazzConcert.Check Poster is not set";
            }

            return true;
        } // Check

    } // JazzConcert
} // namespace
