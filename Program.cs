using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace locate
{
    class Program
    {
        // Connection string for Windows Search
        const string connectionString = "Provider=Search.CollatorDSO;Extended Properties=\"Application=Windows\"";
        static bool OptModified = false;
        static bool OptSize = false;
        static bool OptUsage = false;
        static bool OptPath = true;
        static bool OptQuiet = true;
        static bool OptGroup = false;
        static bool OptRecent = false;

        static void Main(string[] args)
        {
            int chapterDepth = 10;

            string query = null;
            var s0 = args
                .TakeWhile(a=>Regex.IsMatch(a,@"^[-/]"));
            
            var switches = s0
                .Select(a=>Regex.Replace(a,"^[-/]","").ToLower())
                .Aggregate(new StringBuilder(), (sb,w)=>sb.Append(w));

            ProcessArgs(switches);

            args = args.Skip(s0.Count()).ToArray();
            
            try
            {
                //No args: Return top 20 most recently modified files
                if (args.Length == 0)
                {
                    OptRecent = true;
                    query = GetSelectClause();
                    query += " FROM SystemIndex WHERE SCOPE='file:' ORDER BY System.DateModified DESC"; 
                }
                else
                {
                    //We have 1 or more args. We'll treat them as args to a CONTAINS predicate
                    //joined with AND
                    //http://msdn.microsoft.com/en-us/library/bb231270%28v=VS.85%29.aspx
                    var keywords=new string[] {"OR", "AND"};

                    query = GetSelectClause();
                    query += " FROM SystemIndex WHERE SCOPE='file:' AND (";

                    int state = 1;  

                    for (int i = 0; i<args.Length;++i)
                    {
                        var arg = args[i].Replace("'", "''");       //escape ticks
                        arg = args[i].Replace("\"", "\"\"");        //escape quotes
                        
                        arg = "\"" + arg + "\"";                    //enclose in quotes
                        
                        if (i>0 && keywords.Contains(arg))
                        {
                            query += string.Format(" {0} ", arg);
                            state = 1;
                        }
                        else
                        {
                            query += state == 0 ? "AND" : "";
                            query += string.Format(" CONTAINS(System.FileName, '{0}') ",arg);
                            state = 0;
                        }
                    }

                    query += ")";
                    if (OptGroup)
                        query += ")";
                }

                var st = DateTime.Now;
                ExecuteQuery(query, chapterDepth);
                var tt = DateTime.Now - st;
                if (!OptQuiet)
                    Console.WriteLine("Took: {0} ms", tt.TotalMilliseconds);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine();
                Usage();
            }
        }

        private static string GetSelectClause()
        {
            bool bFirstCol = true;
            string query = "";
            
            if (OptRecent)
                query = "SELECT TOP 20 ";
            else if (OptGroup)  //Group && Recent !allowed
                //ORDER BY System.DateModified DESC
                query = "GROUP ON System.ItemPathDisplay OVER (SELECT ";
            else
                query = "SELECT ";


            if (OptModified)
                query += GetOptionalColumn(ref bFirstCol, "System.DateModified");

            if (OptSize)
                query += GetOptionalColumn(ref bFirstCol, "System.Size");
            
            if (OptPath)
                query += GetOptionalColumn(ref bFirstCol, "System.ItemPathDisplay");
            else
                query += GetOptionalColumn(ref bFirstCol, "System.FileName");

            return query;
        }

        private static string GetOptionalColumn(ref bool bFirstCol, string colName)
        {
            string ret = "";
            if (bFirstCol)
            {
                ret = colName;
                bFirstCol = false;
            }
            else
            {
                ret = "," + colName;
            }

            return ret;
        }

        private static void ProcessArgs(StringBuilder switches)
        {
            int state = 0;

            for (int i = 0; i < switches.Length; ++i)
            {
                switch (switches[i])
                {
                    case '-':
                    case '+':
                        break;

                    case 'l':
                        OptModified = (state == 0);
                        OptSize = (state == 0);
                        OptPath = (state == 0);
                        break;

                    case 'm':
                        OptModified = (state == 0);
                        break;

                    case 's':
                        OptSize = (state == 0);
                        break;

                    case 'p':
                        OptPath = (state == 0);
                        break;

                    case 'q':
                        OptQuiet = (state == 0);
                        break;

                    case 'g':
                        OptGroup = (state == 0);
                        break;

                    case 'v':       //verbose (invert quiet)
                        OptQuiet = !(state == 0);
                        break;



                    default:
                        OptUsage = true;
                        break;
                }

                if (OptUsage)
                    Usage(); //and be done with it!

                state = switches[i] == '-' ? 1 : 0;
            }
        }

        // Display the result set recursively expanding chapterDepth deep
        static void DisplayReader(OleDbDataReader myDataReader, ref uint count, uint alignment, int chapterDepth)
        {
            try
            {
                // compute alignment
                StringBuilder indent = new StringBuilder((int)alignment);
                indent.Append(' ', (int)alignment);

                while (myDataReader.Read())
                {
                    // add alignment
                    StringBuilder row = new StringBuilder(indent.ToString());

                    // for all columns
                    for (int i = 0; i < myDataReader.FieldCount; i++)
                    {
                        // null columns
                        if (myDataReader.IsDBNull(i))
                        {
                            row.Append("NULL\t");
                        }
                        else
                        {
                            //vector columns
                            object[] myArray = myDataReader.GetValue(i) as object[];
                            if (myArray != null)
                            {
                                DisplayValue(myArray, row);
                            }
                            else
                            {
                                //check for chapter columns from "group on" queries
                                if (myDataReader.GetFieldType(i).ToString() != "System.Data.IDataReader")
                                {
                                    //regular columns are displayed here
                                    //If (OptPath == false) and colName is System.ItemPathDisplay 
                                    //We want to skip Display of this Column
                                    if (!(OptPath == false && myDataReader.GetName(i) == "System.ItemPathDisplay"))
                                        row.Append(myDataReader.GetValue(i));
                                }
                                else
                                {
                                    //for a chapter column type just display the colum name
                                    row.Append(myDataReader.GetName(i));
                                }
                            }
                            row.Append('\t');
                        }
                    }
                    if (chapterDepth >= 0)
                    {
                        Console.WriteLine(row.ToString());
                        count++;
                    }
                    // for each chapter column
                    for (int i = 0; i < myDataReader.FieldCount; i++)
                    {
                        if (myDataReader.GetFieldType(i).ToString() == "System.Data.IDataReader")
                        {
                            OleDbDataReader Reader = myDataReader.GetValue(i) as OleDbDataReader;
                            DisplayReader(Reader, ref count, alignment + 8, chapterDepth - 1);
                        }
                    }
                }
            }
            finally
            {
                myDataReader.Close();
                myDataReader.Dispose();
            }
        }

        // display the value recursively
        static void DisplayValue(object value, StringBuilder sb)
        {
            if (value != null)
            {
                if (value.GetType().IsArray)
                {
                    sb.Append("[");
                    bool first = true;

                    // display every element
                    foreach (object subval in value as Array)
                    {
                        if (first)
                        {
                            first = false;
                        }
                        else
                        {
                            sb.Append("; ");
                        }
                        DisplayValue(subval, sb);
                    }

                    sb.Append("]");
                }
                else
                {
                    if (value.GetType() == typeof(double))
                    {
                        // Normal numeric formats round, but we want to report the actual round trip format
                        sb.AppendFormat("{0:r}", value);
                    }
                    else
                    {
                        sb.Append(value);
                    }
                }
            }
        }

        // Run a query and display the rowset up to chapterDepth deep
        static void ExecuteQuery(string query, int chapterDepth)
        {
            OleDbDataReader myDataReader = null;
            OleDbConnection myOleDbConnection = new OleDbConnection(connectionString);
            OleDbCommand myOleDbCommand = new OleDbCommand(query, myOleDbConnection);
            try
            {
                if (!OptQuiet)
                    Console.WriteLine("Query=" + query);

                myOleDbConnection.Open();
                myDataReader = myOleDbCommand.ExecuteReader();
                if (!myDataReader.HasRows)
                {
                    System.Console.WriteLine("Query returned 0 rows!");
                    return;
                }
                uint count = 0;
                DisplayReader(myDataReader, ref count, 0, chapterDepth);
                
                if (!OptQuiet)
                    Console.WriteLine("Rows+Chapters=" + count);
            }
            catch (System.Data.OleDb.OleDbException oleDbException)
            {
                Console.WriteLine("Got OleDbException, error code is 0x{0:X}L", oleDbException.ErrorCode);
                Console.WriteLine("Exception details:");
                for (int i = 0; i < oleDbException.Errors.Count; i++)
                {
                    Console.WriteLine("\tError " + i.ToString(CultureInfo.CurrentCulture.NumberFormat) + "\n" +
                                      "\t\tMessage: " + oleDbException.Errors[i].Message + "\n" +
                                      "\t\tNative: " + oleDbException.Errors[i].NativeError.ToString(CultureInfo.CurrentCulture.NumberFormat) + "\n" +
                                      "\t\tSource: " + oleDbException.Errors[i].Source + "\n" +
                                      "\t\tSQL: " + oleDbException.Errors[i].SQLState + "\n");
                }
                Console.WriteLine(oleDbException.ToString());
            }
            finally
            {
                // Always call Close when done reading.
                if (myDataReader != null)
                {
                    myDataReader.Close();
                    myDataReader.Dispose();
                }
                // Close the connection when done with it.
                if (myOleDbConnection.State == System.Data.ConnectionState.Open)
                {
                    myOleDbConnection.Close();
                }
            }
        }

        static void Usage()
        {
            var a = Environment.GetCommandLineArgs()[0];
            Console.WriteLine("Usage: \n");
            Console.WriteLine(a + " [keyword] [AND|OR] [keyword]");
            Console.WriteLine("");
            Console.WriteLine("\nThis will find all files/folders in Windows Search Index where the Filename contains keyword");
            Console.WriteLine("\n");
            Console.WriteLine("    Options: ");
            Console.WriteLine("      -l\t\tLong Displays DateModified and Size of files");
            Console.WriteLine("      -m\t\tShow Date modified");
            Console.WriteLine("      -s\t\tShow Size (NULL for folders)");
            Console.WriteLine("      --\t\tNegate next switch (so --s means do not show size");
            Console.WriteLine("\n");
            Console.WriteLine("Options can be combined, so '-ms' is equivalent to '-m -s' and results in 'Show Date Modified and Size' in results");
                        
            Environment.Exit(0);
        }
    }
}
