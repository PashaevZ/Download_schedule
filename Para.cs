using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace dotnet
{
    class Para
    {
        public DateTime dtm { get; private set; }        
        public string subject { get; private set; }
        public string teacher { get; private set; }
        public string cabinet { get; private set; }

        public Para (DateTime dtm, string subject, string teacher="unknown", string cabinet="000")
        {            
            this.dtm = dtm;
            this.subject = subject;
            this.teacher = teacher;
            this.cabinet = cabinet;
        }
        
        public void Print()
        {
            System.Console.WriteLine("{0} \t{1} \t{2} \t{3}", dtm.ToString(), cabinet, teacher, subject);
        }
    }

    class Rasp 
    {
        public string group { get; private set; }
        public List <Para>  paras { get; private set; }

        public Rasp (string group)
        {
            this.group = group;
            paras = new List<Para>();
        }

        public void AddPara (DateTime dtm, string subject, string teacher, string cabinet)
        {
            Para para = new Para (dtm, subject, teacher, cabinet);
            paras.Add(para);
        }
        public void PrintRasp()
        {
            foreach (Para para in paras)
            {
                para.Print();
            }
        }
        public string SerializeRasp()
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
            };   

            string jsonString = JsonSerializer.Serialize(this, options);
            return jsonString;
        }
        public void SaveAsJson()
        {
            File.WriteAllText("rasp.json", this.SerializeRasp());
            System.Console.WriteLine("Json saved");
        }
    }
}