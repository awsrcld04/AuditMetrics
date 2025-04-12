using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.DirectoryServices;

namespace AuditMetrics
{
    class AMMain
    {
        static DirectorySearcher funcCreateDSSearcher()
        {
            // [Comment] Get local domain context
            string rootDSE;

            System.DirectoryServices.DirectorySearcher objrootDSESearcher = new System.DirectoryServices.DirectorySearcher();
            rootDSE = objrootDSESearcher.SearchRoot.Path;
            // [DebugLine]Console.WriteLine(rootDSE);

            // [Comment] Construct DirectorySearcher object using rootDSE string
            System.DirectoryServices.DirectoryEntry objrootDSEentry = new System.DirectoryServices.DirectoryEntry(rootDSE);
            System.DirectoryServices.DirectorySearcher objDSSearcher = new System.DirectoryServices.DirectorySearcher(objrootDSEentry);
            // [DebugLine]Console.WriteLine(objDSSearcher.SearchRoot.Path);
            return objDSSearcher;
        }

        static void funcLogToEventLog(string strAppName, string strEventMsg, int intEventType)
        {
            string sLog;

            sLog = "Application";

            if (!EventLog.SourceExists(strAppName))
                EventLog.CreateEventSource(strAppName, sLog);

            //EventLog.WriteEntry(strAppName, strEventMsg);
            EventLog.WriteEntry(strAppName, strEventMsg, EventLogEntryType.Information, intEventType);
        } // LogToEventLog

        static void funcPrintParameterSyntax()
        {
            Console.WriteLine("AccountVerifier v1.0 (c) 2011 SystemsAdminPro.com");
            Console.WriteLine();
            Console.WriteLine("Parameter syntax:");
            Console.WriteLine();
            Console.WriteLine("Use the following for the first parameter:");
            Console.WriteLine("[filename]          specify a filename to use as the input file");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("AccountVerifier testinputfile.txt");
        }

        static void funcGetMetrics()
        {
            string strQueryFilterPrefix = "(&(objectCategory=group)(name=";
            string strQueryFilterSuffix = "))";

            string objAccountDEvalues;
            string objAccountNameValue;
            int intStrPosFirst = 3;
            int intStrPosLast;

            string[] arrGroups = {"Domain Admins","Enterprise Admins","Schema Admins"};

            foreach (string strGroupName in arrGroups)
            {
                string strTmpGroupName = strGroupName;

                System.DirectoryServices.DirectorySearcher objAccountObjectSearcher = funcCreateDSSearcher();

                // [Comment] Add filter to DirectorySearcher object
                objAccountObjectSearcher.Filter = (strQueryFilterPrefix + strTmpGroupName + strQueryFilterSuffix);

                // [Comment] Execute query, return results, display name and path values
                System.DirectoryServices.SearchResultCollection objAccountResults = objAccountObjectSearcher.FindAll();

                foreach (System.DirectoryServices.SearchResult objAccount in objAccountResults)
                {
                    System.DirectoryServices.DirectoryEntry objAccountDE = new System.DirectoryServices.DirectoryEntry(objAccount.Path);
                    intStrPosLast = objAccountDE.Name.Length;
                    objAccountNameValue = objAccountDE.Name.Substring(intStrPosFirst, intStrPosLast - intStrPosFirst);

                    objAccountDEvalues = objAccountNameValue + "\t" + objAccountDE.Path;
                    Console.WriteLine("AD group exists for: {0}", objAccountDEvalues);

                    funcGetGroupMembers(objAccountDE.Path);
                }
            }

            Console.WriteLine();
        }

        static void funcGetGroupMembers(string strGroupPath)
        {
            string strQueryFilterPrefix = "(&(objectCategory=user)(memberOf=";
            string strQueryFilterSuffix = "))";

            string objAccountDEvalues;
            string objAccountNameValue;
            int intStrPosFirst = 3;
            int intStrPosLast;

            string strTmpGroupName = strGroupPath.Substring(7,strGroupPath.Length-7);

            System.DirectoryServices.DirectorySearcher objAccountObjectSearcher = funcCreateDSSearcher();

            // [Comment] Add filter to DirectorySearcher object
            objAccountObjectSearcher.Filter = (strQueryFilterPrefix + strTmpGroupName + strQueryFilterSuffix);

            // [Comment] Execute query, return results, display name and path values
            System.DirectoryServices.SearchResultCollection objAccountResults = objAccountObjectSearcher.FindAll();

            foreach (System.DirectoryServices.SearchResult objAccount in objAccountResults)
            {
                System.DirectoryServices.DirectoryEntry objAccountDE = new System.DirectoryServices.DirectoryEntry(objAccount.Path);
                intStrPosLast = objAccountDE.Name.Length;
                objAccountNameValue = objAccountDE.Name.Substring(intStrPosFirst, intStrPosLast - intStrPosFirst);

                objAccountDEvalues = objAccountNameValue + "\t" + objAccountDE.Path;
                Console.WriteLine("Group member: {0}", objAccountDEvalues);
            }

        }

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Parameters must be specified to run AuditMetrics.");
                Console.WriteLine("Run AuditMetrics -? to get the parameter syntax.");
            }
            else
            {
                if (args[0] == "-?")
                {
                    funcPrintParameterSyntax();
                }
                else
                {
                    funcGetMetrics();
                }
            }
        }
    }
}
