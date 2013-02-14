using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommandLine.Utility;

namespace BDOC.MigrationTools
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Console.WriteLine("BDOC Style Updater. ");
           
            Arguments CommandLine = new Arguments(args);

            if (CommandLine["h"] != null)
            {
                System.Console.WriteLine("Parameters: \n -dir DirectoryOfFiles : Files directory.");
                System.Console.WriteLine("-l  : List all doc files with TAB character.");
                System.Console.WriteLine("-f  : List all font and styles in doc files.");
                System.Console.WriteLine("-m  : Modify all Styles in doc files.");
                System.Console.WriteLine("-t  : modify all left indents and autoSize in tables for all doc files.");
                System.Console.WriteLine("-tw  : modify teh PreferredWidthType to pourcentage for all tables in all doc files.");
                System.Console.WriteLine("-e type : modify all Inter-line in doc Files with type policy: 1=LineSpaceSingle,2=LineSpace1pt5,3=LineSpaceDouble,4=LineSpaceatLeast,5=LineSpaceExactly,6=LineSpaceMultiple.");
                System.Console.WriteLine("-i number : Inter-line lineSpace number (necessary for types 4,5,6, if not given default = 10, not necessary for 1,2,3).");               
            }
            else
            {
                if (CommandLine["dir"] == null)
                {
                    System.Console.WriteLine("ERROR: Directory Not Found. Please use -dir=DirectoryOfFiles parameter.");
                    return;
                }


                try
                {

                    RenameWordStyles renObj = new RenameWordStyles();
                    renObj.startModification(args);
                }

                catch (Exception ex)
                {
                    System.Console.WriteLine("Error Found. " + ex.Message);
                }
            }

            
        }
    }
}
