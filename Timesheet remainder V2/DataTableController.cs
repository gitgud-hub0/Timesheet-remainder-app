using System;
using System.Collections.Generic;
using System.Data;

namespace Timesheet_remainder
{
    public class DataTableController
    {
        public DataTable GetOutputDataTable(List<object> instanceList)
        {
            var inputTable = PopulateInputTable(instanceList);

            var sortCalcTable = SortCalcTable(inputTable);
            sortCalcTable.WriteXml(@"C:\Users\Ben\Desktop\Timesheet test\calcTest\calcTable.xml");

            var outTable = PopulateOutputTable(sortCalcTable);
            outTable.WriteXml(@"C:\Users\Ben\Desktop\Timesheet test\calcTest\outTable.xml");

            return outTable;
        }

        private DataTable PopulateInputTable(List<object> instanceList)
        {
            DataTable inputTable = new DataTable("inputTable1");
            inputTable.Clear();
            inputTable.Columns.Add("Date");
            inputTable.Columns.Add("Time");
            inputTable.Columns.Add("TaskDesc");
            for (int order = 0; order <= instanceList.Count - 2; order += 3)
            {
                DataRow task = inputTable.NewRow();
                task["Date"] = Convert.ToString(instanceList[order]);
                task["Time"] = Convert.ToString(instanceList[order + 1]);
                task["TaskDesc"] = Convert.ToString(instanceList[order + 2]);
                inputTable.Rows.Add(task);
            }

            return inputTable;
        }

        private DataTable SortCalcTable(DataTable inputTable)
        {
            var calcTable = new DataTable("calcTable1");
            calcTable.Columns.Add("Date");
            calcTable.Columns.Add("Task");
            calcTable.Columns.Add("Duration");
            for (int i = 0; i < inputTable.Rows.Count - 1; i++)
            {
                // Compare with previous row using index
                if (inputTable.Rows[i]["Date"] == inputTable.Rows[i + 1]["Date"])
                {
                    var calcRow = calcTable.NewRow();
                    calcRow["Date"] = inputTable.Rows[i]["Date"];
                    calcRow["Task"] = inputTable.Rows[i]["TaskDesc"];
                    calcRow["Duration"] =
                        (DateTime.Parse((string) inputTable.Rows[i + 1]["Time"])
                            .Subtract(DateTime.Parse((string) inputTable.Rows[i]["Time"]))
                        ).TotalMinutes;
                    calcTable.Rows.Add(calcRow);
                }
            }

            return SortDataTable(calcTable);
        }

        private DataTable SortDataTable(DataTable calcTable)
        {
            //sort table in terms of date then tasks using linq
            var sortedCalcTable = calcTable.AsEnumerable()
                .OrderBy(r => r.Field<string>("Date"))
                .ThenBy(r => r.Field<string>("Task"))
                .CopyToDataTable();
            sortedCalcTable.TableName = "SortedCalcTable";
            return sortedCalcTable;
        }

        private DataTable PopulateOutputTable(DataTable sortCalcTable)
        {
            //output table
            DataTable outTable = new DataTable("OutputTable1");
            outTable.Clear();
            outTable.Columns.Add("Date");
            outTable.Columns.Add("Task");
            outTable.Columns.Add("TotalTime");
            outTable.Columns.Add("TotalDuration", typeof(int));

            //generate rows with unique tasks each 
            DataRow outRowInitial = outTable.NewRow();
            outRowInitial["Date"] = sortCalcTable.Rows[0]["Date"];
            outRowInitial["Task"] = sortCalcTable.Rows[0]["Task"];
            outRowInitial["TotalDuration"] = 0;
            outTable.Rows.Add(outRowInitial);
            //int j = 0; //iterator for outtable
            for (int i = 0; i < sortCalcTable.Rows.Count; i++) //iterator for sortCalctable
            {
                if (sortCalcTable.Rows[i]["Date"] == outTable.Rows[outTable.Rows.Count - 1]["Date"])
                {
                    if (sortCalcTable.Rows[i]["Task"] != outTable.Rows[outTable.Rows.Count - 1]["Task"])
                    {
                        DataRow outRow = outTable.NewRow();
                        outRow["Date"] = sortCalcTable.Rows[i]["Date"];
                        outRow["Task"] = sortCalcTable.Rows[i]["Task"];
                        outRow["TotalDuration"] = 0;
                        outRow["TotalTime"] = 0;
                        outTable.Rows.Add(outRow);
                    }
                }
                else
                {
                    DataRow outRow = outTable.NewRow();
                    outRow["Date"] = sortCalcTable.Rows[i]["Date"];
                    outRow["Task"] = sortCalcTable.Rows[i]["Task"];
                    outRow["TotalDuration"] = 0;
                    outRow["TotalTime"] = 0;
                    outTable.Rows.Add(outRow);
                }

                if (sortCalcTable.Rows[i]["Date"] == outTable.Rows[outTable.Rows.Count - 1]["Date"]
                    && sortCalcTable.Rows[i]["Task"] == outTable.Rows[outTable.Rows.Count - 1]["Task"])
                {
                    var sum = Convert.ToInt32(outTable.Rows[outTable.Rows.Count - 1]["TotalDuration"])
                              + Convert.ToInt32(sortCalcTable.Rows[i]["Duration"]);
                    var ts = TimeSpan.FromMinutes(sum);
                    DataRow outRow = outTable.Rows[outTable.Rows.Count - 1];
                    outRow["TotalDuration"] = sum;
                    outRow["TotalTime"] = ts.ToString("hh\\:mm");
                }
            }

            return outTable;
        }
    }
}