using System;
using System.Collections.Generic;
using System.Data;

namespace Timesheet_remainder
{
    public class DataTableController
    {
        public DataTableController()
        {
        }

        public DataTable PopulateInputTable(List<object> instanceList)
        {
            DataTable inputTable = new DataTable("inputTable1");
            inputTable.Clear();
            inputTable.Columns.Add("Date");
            inputTable.Columns.Add("Time");
            inputTable.Columns.Add("TaskDesc");
            for (int order = 0; order <= instanceList.Count - 2; order += 3)
            {
                DataRow _task = inputTable.NewRow();
                _task["Date"] = Convert.ToString(instanceList[order]);
                _task["Time"] = Convert.ToString(instanceList[order + 1]);
                _task["TaskDesc"] = Convert.ToString(instanceList[order + 2]);
                inputTable.Rows.Add(_task);
            }

            return inputTable;
        }

        public DataTable SortCalcTable(DataTable inputTable)
        {
            DataTable calcTable = new DataTable("calcTable1");
            calcTable.Clear();
            calcTable.Columns.Add("Date");
            calcTable.Columns.Add("Task");
            calcTable.Columns.Add("Duration");
            for (int i = 0; i < inputTable.Rows.Count - 1; i++)
            {
                // Compare with previous row using index
                if (inputTable.Rows[i]["Date"] == inputTable.Rows[i + 1]["Date"])
                {
                    DataRow _calcRow = calcTable.NewRow();
                    _calcRow["Date"] = inputTable.Rows[i]["Date"];
                    _calcRow["Task"] = inputTable.Rows[i]["TaskDesc"];
                    _calcRow["Duration"] =
                        (DateTime.Parse((string) inputTable.Rows[i + 1]["Time"])
                            .Subtract(DateTime.Parse((string) inputTable.Rows[i]["Time"]))).TotalMinutes;
                    calcTable.Rows.Add(_calcRow);
                }
            }

            //sort table in terms of date then tasks using linq
            DataTable sortCalcTable = new DataTable();
            sortCalcTable.Clear();
            sortCalcTable = calcTable.AsEnumerable()
                .OrderBy(r => r.Field<string>("Date"))
                .ThenBy(r => r.Field<string>("Task"))
                .CopyToDataTable();
            sortCalcTable.TableName = "sortCalcTable1";
            return sortCalcTable;
        }

        public DataTable PopulateOutputTable(DataTable sortCalcTable)
        {
            //output table
            DataTable outTable = new DataTable("OutputTable1");
            outTable.Clear();
            outTable.Columns.Add("Date");
            outTable.Columns.Add("Task");
            outTable.Columns.Add("TotalTime");
            outTable.Columns.Add("TotalDuration", typeof(int));

            //generate rows with unique tasks each 
            DataRow _outRow1st = outTable.NewRow();
            _outRow1st["Date"] = sortCalcTable.Rows[0]["Date"];
            _outRow1st["Task"] = sortCalcTable.Rows[0]["Task"];
            _outRow1st["TotalDuration"] = 0;
            outTable.Rows.Add(_outRow1st);
            //int j = 0; //iterator for outtable
            int sum = 0;
            TimeSpan ts;
            for (int i = 0; i < sortCalcTable.Rows.Count; i++) //iterator for sortCalctable
            {
                if (sortCalcTable.Rows[i]["Date"] == outTable.Rows[outTable.Rows.Count - 1]["Date"])
                {
                    if (sortCalcTable.Rows[i]["Task"] != outTable.Rows[outTable.Rows.Count - 1]["Task"])
                    {
                        DataRow _outRow = outTable.NewRow();
                        _outRow["Date"] = sortCalcTable.Rows[i]["Date"];
                        _outRow["Task"] = sortCalcTable.Rows[i]["Task"];
                        _outRow["TotalDuration"] = 0;
                        _outRow["TotalTime"] = 0;
                        outTable.Rows.Add(_outRow);
                    }
                }
                else
                {
                    DataRow _outRow = outTable.NewRow();
                    _outRow["Date"] = sortCalcTable.Rows[i]["Date"];
                    _outRow["Task"] = sortCalcTable.Rows[i]["Task"];
                    _outRow["TotalDuration"] = 0;
                    _outRow["TotalTime"] = 0;
                    outTable.Rows.Add(_outRow);
                }

                if (sortCalcTable.Rows[i]["Date"] == outTable.Rows[outTable.Rows.Count - 1]["Date"]
                    && sortCalcTable.Rows[i]["Task"] == outTable.Rows[outTable.Rows.Count - 1]["Task"])
                {
                    sum = Convert.ToInt32(outTable.Rows[outTable.Rows.Count - 1]["TotalDuration"])
                          + Convert.ToInt32(sortCalcTable.Rows[i]["Duration"]);
                    ts = TimeSpan.FromMinutes(sum);
                    DataRow _outRow = outTable.Rows[outTable.Rows.Count - 1];
                    _outRow["TotalDuration"] = sum;
                    _outRow["TotalTime"] = ts.ToString("hh\\:mm");

                    sum = 0;
                }
            }

            return outTable;
        }
    }
}