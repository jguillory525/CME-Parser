using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;

namespace CME_Parser
{
    
    public partial class Form1 : Form
    {
        public string[] Counties = {
       "0-Alamance",    "1-Alexander",  "2-Alleghany",  "3-Anson",  "4-Ashe",   "5-Avery ", "6-Beaufort",
       "7-Bertie",  "8-Bladen", "9-Brunswick",  "10-Buncombe",  "11-Burke", "12-Cabarrus",  "13-Caldwell",
       "14-Camden", "15-Carteret",  "16-Caswell",   "17-Catawaba",  "18-Chatham",   "19-Cherokee",  "20-Chowan",
       "21-Clay",   "22-Cleveland", "23-Columbus",  "24-Craven",    "25-Cumberland",    "26-Currituck", "27-Dare",
       "28-Davidson",   "29-Davie ",    "30-Duplin ",   "31-Durham",    "32-Edgecomb",  "33-Forsyth",
       "34-Franklin ", "35-Gaston",    "36-Gates ",    "37-Graham",    "38-Granville", "39-Greene",    "40-Guilford",
       "41-Halifax ",  "42-Harnett",   "43-Haywood",   "44-Henderson", "45-Hertford",  "46-Hoke",  "47-Hyde",
       "48-Iredell",   "49-Jackson",   "50-Johnston",  "51-Jones", "52-Lee",   "53-Lenoir",    "54-Lincoln",   "55-Macon",
       "56-Madison",   "57-Martin",    "58-McDowell",  "59-Mecklenburg",   "60-Mitchell",  "61-Montgomery",    "62-Moore",
       "63-Nash",  "64-New Hanover",   "65-Northampton",   "66-Onslow",    "67-Orange", "68-Pamlico",   "69-Pasquotank",
       "70-Pender",    "71-Perquimans",    "72-Person",    "73-Pitt",  "74-Polk",  "75-Randolph",  "76-Richmond",
       "77-Robeson",   "78-Rockingham",    "79-Rowan", "80-Rutherford",    "81-Sampson",   "82-Scotland",  "83-Stanley",
       "84-Stokes",    "85-Surry", "86-Swain", "87-Transylvania",  "88-Tyrrell",   "89-Union", "90-Vance", "91-Wake",
       "92-Warren",    "93-Washington",    "94-Watauga",   "95-Wayne", "96-Wilkes",    "97-Wilson",    "98-Yadkin",
       "99-Yancey"
        };

        public class DN 
        {
            public string id;
            public string name;
            public string number;
            public string label;
            public string pickupGroup;
            public string fwdAll;
            public string fwdBusy;
            public string fwdNoan;
            public string fwdUnreg;
        
        }
        public class BLF
        {
            public string number;
            public string label;
        }

        public class SpeedDial
        {
            public string number;
            public string label;
        }
        public class Pool
        {
            public string id;
            public string mac;
            public string description;
            public string type;
            public DN dn1;
            public DN dn2;
            public DN dn3;
            public DN dn4;
            public DN dn5;
            public string epnm1;
            public string epnm2;
            public string epnm3;
            public string epnm4;
            public string epnm5;
            public string cor1;
            public string cor2;
            public string cor3;
            public string cor4;
            public string cor5;
            public BLF[] blfs;
            public SpeedDial[] speedDials;
            public string template;
        
        }

        public class data
        {
            public string filePath;
            public string siteNumber;
            public string siteName;
            public List<DN> dns;
            public List<Pool> pools;
        }

        public data parse;
        public Form1()
        {
            
            InitializeComponent();
            foreach (string county in Counties)
            {
                comboBox1.Items.Add(county);
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                parse = new data();
                parse.dns = new List<DN>();
                parse.pools = new List<Pool>();
                
                //Get the path of specified file
                parse.filePath = openFileDialog1.FileName;
                label1.Visible = true;
                comboBox1.Enabled = true;
                prefixBox.Enabled = true;
                button4.Enabled = true;
               
            }
        }

        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
         
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StreamReader reader = File.OpenText(parse.filePath);
            string line;
            string[] words;
            while ((line = reader.ReadLine()) != null)
            {

                if (line.StartsWith("voice register dn  "))
                {
                    string prefix = prefixBox.Text;
                    DN currentDn = new DN();
                    words = line.Split(' ');
                    currentDn.id = words[4];
                    textBox.AppendText(currentDn.id + Environment.NewLine);

                    while ((line = reader.ReadLine()) != "!")
                    {  
                        if (line.StartsWith(" number"))
                        {
                            words = line.Split(' ');
                            currentDn.number = prefix + words[2];
                            textBox.AppendText(currentDn.number + Environment.NewLine);


                        }
                        if (line.StartsWith(" name"))
                        {
                            currentDn.name = line.Remove(0, 6);
                            textBox.AppendText(currentDn.name + Environment.NewLine);


                        }
                        if (line.StartsWith(" label"))
                        {
                            currentDn.label = line.Remove(0, 7);
                            textBox.AppendText(currentDn.label + Environment.NewLine);


                        }
                        if (line.StartsWith(" pickup-group"))
                        {
                            words = line.Split(' ');
                            currentDn.pickupGroup = parse.siteName + words[2];
                            textBox.AppendText("pickup:"+currentDn.pickupGroup + Environment.NewLine);


                        }
                        if (line.StartsWith(" call-forward b2bua all"))
                        {
                            words = line.Split(' ');
                            currentDn.fwdAll = prefix + words[4];
                            textBox.AppendText("all:"+currentDn.fwdAll + Environment.NewLine);


                        }
                        if (line.StartsWith(" call-forward b2bua unregistered"))
                        {
                            words = line.Split(' ');
                            currentDn.fwdUnreg = prefix + words[4];
                            textBox.AppendText("uregistered:"+currentDn.fwdUnreg + Environment.NewLine);


                        }
                        if (line.StartsWith(" call-forward b2bua busy"))
                        {
                            words = line.Split(' ');
                            currentDn.fwdBusy = prefix + words[4];
                            textBox.AppendText("busy:"+ currentDn.fwdBusy + Environment.NewLine);


                        }
                        if (line.StartsWith(" call-forward b2bua noan"))
                        {
                            words = line.Split(' ');
                            currentDn.fwdNoan = prefix + words[4];
                            textBox.AppendText("noan:"+ currentDn.fwdNoan + Environment.NewLine);


                        }
                    }
                    parse.dns.Add(currentDn);
                }
            }
            textBox.AppendText("EOF");
            reader.Close();
            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string prefix = prefixBox.Text;
            StreamReader reader = File.OpenText(parse.filePath);
            string line;
            string[] words;
            while ((line = reader.ReadLine()) != null)
            {
                if (line.StartsWith("voice register pool  "))
                {
                    Pool currentPool = new Pool();
                    currentPool.blfs = new BLF[31];
                    currentPool.speedDials = new SpeedDial[31];
                    for (int i = 1; i < 31; i++)
                    {
                        currentPool.blfs[i] = new BLF();
                        currentPool.blfs[i].label = "";
                        currentPool.blfs[i].number = "";
                        currentPool.speedDials[i] = new SpeedDial();
                        currentPool.speedDials[i].label = "";
                        currentPool.speedDials[i].number = "";
                    }

                        words = line.Split(' ');
                    currentPool.id = words[4];
                    textBox.AppendText(currentPool.id + Environment.NewLine);

                    while ((line = reader.ReadLine()) != "!")
                    {
                        if (line.StartsWith(" blf-speed-dial"))
                        {
                            words = line.Split(' ');
                            int index = Convert.ToInt32(words[2]);
                            if (words[3].StartsWith("9"))
                            { 
                                currentPool.blfs[index].number = words[3]; 
                            }
                            else 
                            {
                                currentPool.blfs[index].number = "\\" + prefix + words[3];
                            }
                            
                            //textBox.AppendText(words[2] + " " + words[3]);
                            words = line.Split('"');
                            currentPool.blfs[index].label = words[1];
                            //textBox.AppendText(words[1]);
                        }
                        if (line.StartsWith(" speed-dial"))
                        {
                            words = line.Split(' ');
                            int index = Convert.ToInt32(words[2]);
                            if (words[3].StartsWith("9"))
                            {
                                currentPool.speedDials[index].number = words[3];
                            }
                            else
                            {
                                currentPool.speedDials[index].number = prefix + words[3];
                            }
                      
                            //textBox.AppendText(words[2] + " " + words[3]);
                            words = line.Split('"');
                            currentPool.speedDials[index].label = words[1];
                            //textBox.AppendText(words[1]);
                        }


                        if (line.StartsWith(" id mac"))
                        {
                            words = line.Split(' ');
                            currentPool.mac = "SEP" + (words[3].Replace(".",string.Empty)).ToUpper();
                            textBox.AppendText(currentPool.mac + Environment.NewLine);

                        }

                        if (line.StartsWith(" type"))
                        {
                            words = line.Split(' ');
                            currentPool.type = "Cisco "+words[2];
                            textBox.AppendText(currentPool.type + Environment.NewLine);

                        }
                     /* pool description not being used, use the name of DN 1 prefixed by site code
                        if (line.StartsWith(" description"))
                        {
                            words = line.Split(' ');
                            if (currentPool.description == null)
                                currentPool.description = words[2];
                            textBox.AppendText(currentPool.description + Environment.NewLine);

                        }
                        */
                        if (line.StartsWith(" number 1"))
                        {
                            words = line.Split(' ');


                            currentPool.dn1 = (from dns in parse.dns
                                             where dns.id == words[4]
                                             select dns).FirstOrDefault();

                            currentPool.description = parse.siteNumber + "-" + currentPool.dn1.name;
                            
                           // currentPool.dn1 = prefix + words[4];
                            textBox.AppendText("N1:"+currentPool.dn1.number + Environment.NewLine);

                        }
                        if (line.StartsWith(" number 2"))
                        {
                            words = line.Split(' ');
                            currentPool.dn2 = (from dns in parse.dns
                                               where dns.id == words[4]
                                               select dns).FirstOrDefault();
                            //  currentPool.number2 = prefix + words[4];
                            textBox.AppendText("N2:" + currentPool.dn2.number + Environment.NewLine);

                        }
                        if (line.StartsWith(" number 3"))
                        {
                            
                            words = line.Split(' ');
                            currentPool.dn3 = (from dns in parse.dns
                                               where dns.id == words[4]
                                               select dns).FirstOrDefault();
                            // currentPool.number3 = prefix + words[4];
                            textBox.AppendText("N3:" + currentPool.dn3.number + Environment.NewLine);

                        }
                        if (line.StartsWith(" number 4"))
                        {
                            
                            words = line.Split(' ');
                            currentPool.dn4 = (from dns in parse.dns
                                               where dns.id == words[4]
                                               select dns).FirstOrDefault();
                            // currentPool.number4 = prefix + words[4];
                            textBox.AppendText("N4:" + currentPool.dn4.number + Environment.NewLine);

                        }
                        if (line.StartsWith(" number 5"))
                        {
                           
                            words = line.Split(' ');
                            currentPool.dn5 = (from dns in parse.dns
                                               where dns.id == words[4]
                                               select dns).FirstOrDefault();
                            //  currentPool.number5 = prefix + words[4];
                            textBox.AppendText("N5:" + currentPool.dn5.number + Environment.NewLine);

                        }
                        if (line.StartsWith(" template"))
                        {
                            words = line.Split(' ');
                            currentPool.template = words[2];
                            textBox.AppendText("Template:" + currentPool.template + Environment.NewLine);

                        }
                        if (line.StartsWith(" cor incoming"))
                        {
                            words = line.Split(' ');
                            string epnm;


                            switch (words[4])
                            {
                                case "1":
                                    currentPool.cor1 = words[3].Split('-')[1];


                                  
                                    switch (words[3])
                                    {
                                        case "ch-lc-css":
                                            epnm = currentPool.dn1.number;
                                            break;
                                        case "ch-ld-css":
                                            epnm = currentPool.dn1.number;
                                            break;
                                        case "ch-intl-css":
                                            epnm = currentPool.dn1.number;
                                            break;
                                        default:
                                            epnm = prefix + words[3].Substring(words[3].Length - 8, 4);
                                            break;
                                    }


                                    currentPool.epnm1 = epnm;
                                    textBox.AppendText("EPNM1:" + currentPool.epnm1 + Environment.NewLine);
                                    textBox.AppendText("COR1:" + currentPool.cor1 + Environment.NewLine);
                                    break;
                                case "2":


                                    currentPool.cor2 = words[3].Split('-')[1];

                                    switch (words[3])
                                    {
                                        case "ch-lc-css":
                                            epnm = currentPool.dn2.number;
                                            break;
                                        case "ch-ld-css":
                                            epnm = currentPool.dn2.number;
                                            break;
                                        case "ch-intl-css":
                                            epnm = currentPool.dn2.number;
                                            break;
                                        default:
                                            epnm = prefix + words[3].Substring(words[3].Length - 8, 4);
                                            break;
                                    }


                                    currentPool.epnm2 = epnm;
                                    
                                    textBox.AppendText("EPNM2:" + currentPool.epnm2 + Environment.NewLine);
                                    textBox.AppendText("COR2:" + currentPool.cor2 + Environment.NewLine);
                                    break;
                                case "3":
                                    currentPool.cor3 = words[3].Split('-')[1];


                                    switch (words[3])
                                    {
                                        case "ch-lc-css":
                                            epnm = currentPool.dn3.number;
                                            break;
                                        case "ch-ld-css":
                                            epnm = currentPool.dn3.number;
                                            break;
                                        case "ch-intl-css":
                                            epnm = currentPool.dn3.number;
                                            break;
                                        default:
                                            epnm = prefix + words[3].Substring(words[3].Length - 8, 4);
                                            break;
                                    }


                                    currentPool.epnm3 = epnm;
                                    textBox.AppendText("EPNM3:" + currentPool.epnm3 + Environment.NewLine);
                                    textBox.AppendText("COR3:" + currentPool.cor3 + Environment.NewLine);
                                    break;
                                case "4":
                                    currentPool.cor4 = words[3].Split('-')[1];


                                    switch (words[3])
                                    {
                                        case "ch-lc-css":
                                            epnm = currentPool.dn4.number;
                                            break;
                                        case "ch-ld-css":
                                            epnm = currentPool.dn4.number;
                                            break;
                                        case "ch-intl-css":
                                            epnm = currentPool.dn4.number;
                                            break;
                                        default:
                                            epnm = prefix + words[3].Substring(words[3].Length - 8, 4);
                                            break;
                                    }


                                    currentPool.epnm4 = epnm;
                                    textBox.AppendText("EPNM4:" + currentPool.epnm4 + Environment.NewLine);
                                    textBox.AppendText("COR4:" + currentPool.cor4 + Environment.NewLine);
                                    break;
                                case "5":
                                    currentPool.cor5 = words[3].Split('-')[1];

                                    switch (words[3])
                                    {
                                        case "ch-lc-css":
                                            epnm = currentPool.dn5.number;
                                            break;
                                        case "ch-ld-css":
                                            epnm = currentPool.dn5.number;
                                            break;
                                        case "ch-intl-css":
                                            epnm = currentPool.dn5.number;
                                            break;
                                        default:
                                            epnm = prefix + words[3].Substring(words[3].Length - 8, 4);
                                            break;
                                    }


                                    currentPool.epnm5 = epnm;
                                    textBox.AppendText("EPNM5:" + currentPool.epnm5 + Environment.NewLine);
                                    textBox.AppendText("COR5:" + currentPool.cor5 + Environment.NewLine);
                                    break;

                            }
                          

                        }

                    }
                    parse.pools.Add(currentPool);
                }
            }
            textBox.AppendText("EOF");
            reader.Close();
            button2.Enabled = false;
            button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            var export = new System.Data.DataTable("Export");

            export.Columns.Add("MAC Address", typeof(string));
            export.Columns.Add("Device Type", typeof(string));
            export.Columns.Add("Template", typeof(string));

            export.Columns.Add("Description", typeof(string));

            export.Columns.Add("Directory Number 1", typeof(string));
            export.Columns.Add("Line CSS 1", typeof(string));
            export.Columns.Add("Pickup Group 1", typeof(string));
            export.Columns.Add("Line Text Label 1", typeof(string));
            export.Columns.Add("External Phone Number Mask 1", typeof(string));
            export.Columns.Add("Alerting Name 1", typeof(string));
            export.Columns.Add("Line Description 1", typeof(string));
            export.Columns.Add("ASCI Alerting Name 1", typeof(string));
            export.Columns.Add("Display 1", typeof(string));
            export.Columns.Add("ASCII Display 1", typeof(string));


            export.Columns.Add("Directory Number 2", typeof(string));
            export.Columns.Add("Line CSS 2", typeof(string));
            export.Columns.Add("Pickup Group 2", typeof(string));
            export.Columns.Add("Line Text Label 2", typeof(string));
            export.Columns.Add("External Phone Number Mask 2", typeof(string));
            export.Columns.Add("Alerting Name 2", typeof(string));
            export.Columns.Add("Line Description 2", typeof(string));
            export.Columns.Add("ASCI Alerting Name 2", typeof(string));
            export.Columns.Add("Display 2", typeof(string));
            export.Columns.Add("ASCII Display 2", typeof(string));


            export.Columns.Add("Directory Number 3", typeof(string));
            export.Columns.Add("Line CSS 3", typeof(string));
            export.Columns.Add("Pickup Group 3", typeof(string));
            export.Columns.Add("Line Text Label 3", typeof(string));
            export.Columns.Add("External Phone Number Mask 3", typeof(string));
            export.Columns.Add("Alerting Name 3", typeof(string));
            export.Columns.Add("Line Description 3", typeof(string));
            export.Columns.Add("ASCI Alerting Name 3", typeof(string));
            export.Columns.Add("Display 3", typeof(string));
            export.Columns.Add("ASCII Display 3", typeof(string));

            export.Columns.Add("Directory Number 4", typeof(string));
            export.Columns.Add("Line CSS 4", typeof(string));
            export.Columns.Add("Pickup Group 4", typeof(string));
            export.Columns.Add("Line Text Label 4", typeof(string));
            export.Columns.Add("External Phone Number Mask 4", typeof(string));
            export.Columns.Add("Alerting Name 4", typeof(string));
            export.Columns.Add("Line Description 4", typeof(string));
            export.Columns.Add("ASCI Alerting Name 4", typeof(string));
            export.Columns.Add("Display 4", typeof(string));
            export.Columns.Add("ASCII Display 4", typeof(string));


            export.Columns.Add("Directory Number 5", typeof(string));
            export.Columns.Add("Line CSS 5", typeof(string));
            export.Columns.Add("Pickup Group 5", typeof(string));
            export.Columns.Add("Line Text Label 5", typeof(string));
            export.Columns.Add("External Phone Number Mask 5", typeof(string));
            export.Columns.Add("Alerting Name 5", typeof(string));
            export.Columns.Add("Line Description 5", typeof(string));
            export.Columns.Add("ASCI Alerting Name 5", typeof(string));
            export.Columns.Add("Display 5", typeof(string));
            export.Columns.Add("ASCII Display 5", typeof(string));

            for (int i = 1; i<31; i++)
            {
                export.Columns.Add("Busy Lamp Field Destination " + i.ToString(), typeof(string)); ;
                export.Columns.Add("Busy Lamp Field Label " + i.ToString(), typeof(string));
            }
            for (int i = 1; i < 31; i++)
            {
                export.Columns.Add("Speed Dial Number " + i.ToString(), typeof(string)); ;
                export.Columns.Add("Speed Dial Label " + i.ToString(), typeof(string));
            }





            foreach (Pool pool in parse.pools)
            {
                string dn1, css1, pg1, ltl1, epnm1, an1, ld1, aan1, d1, ad1,
                    dn2, css2, pg2, ltl2, epnm2, an2, ld2, aan2, d2, ad2,
                    dn3, css3, pg3, ltl3, epnm3, an3, ld3, aan3, d3, ad3,
                    dn4, css4, pg4, ltl4, epnm4, an4, ld4, aan4, d4, ad4,
                    dn5, css5, pg5, ltl5, epnm5, an5, ld5, aan5, d5, ad5;

                dn1=  css1=  pg1= ltl1= epnm1= an1= ld1= aan1= d1= ad1=
                    dn2= css2= pg2= ltl2= epnm2= an2= ld2= aan2= d2= ad2=
                    dn3= css3= pg3= ltl3= epnm3= an3= ld3= aan3= d3= ad3=
                    dn4= css4= pg4= ltl4= epnm4= an4= ld4= aan4= d4= ad4=
                    dn5= css5= pg5= ltl5= epnm5= an5= ld5= aan5= d5= ad5 = "";

                if (!(pool.dn1 is null))
                {
                    dn1 = "\\" + pool.dn1.number;
                    css1 = pool.cor1;
                    pg1 = pool.dn1.pickupGroup;
                    ltl1 = pool.dn1.label;
                    epnm1 = pool.epnm1;
                    an1 = ld1 = aan1 = d1 =ad1 = pool.dn1.name;           
                
                }

                if (!(pool.dn2 is null))
                {
                    dn2 = "\\" + pool.dn2.number;
                    css2 = pool.cor2;
                    pg2 = pool.dn2.pickupGroup;
                    ltl2 = pool.dn2.label;
                    epnm2 = pool.epnm2;
                    an2 = ld2 = aan2 = d2 = ad2 = pool.dn2.name;

                }

                if (!(pool.dn3 is null))
                {
                    dn3 = "\\" + pool.dn3.number;
                    css3 = pool.cor3;
                    pg3 = pool.dn3.pickupGroup;
                    ltl3 = pool.dn3.label;
                    epnm3 = pool.epnm3;
                    an3 = ld3 = aan3 = d3 = ad3 = pool.dn3.name;

                }
                if (!(pool.dn4 is null))
                {
                    dn4 = "\\" + pool.dn4.number;
                    css4 = pool.cor4;
                    pg4 = pool.dn4.pickupGroup;
                    ltl4 = pool.dn4.label;
                    epnm4 = pool.epnm4;
                    an4 = ld4 = aan4 = d4 = ad4 = pool.dn4.name;

                }
                if (!(pool.dn5 is null))
                {
                    dn5 = "\\" + pool.dn5.number;
                    css5 = pool.cor5;
                    pg5 = pool.dn5.pickupGroup;
                    ltl5 = pool.dn5.label;
                    epnm5 = pool.epnm5;
                    an5 = ld5 = aan5 = d5 = ad5 = pool.dn5.name;

                }

                export.Rows.Add(pool.mac, pool.type, pool.template, pool.description, dn1, css1, pg1, ltl1, epnm1, an1, ld1, aan1, d1, ad1,
                    dn2, css2, pg2, ltl2, epnm2, an2, ld2, aan2, d2, ad2,
                    dn3, css3, pg3, ltl3, epnm3, an3, ld3, aan3, d3, ad3,
                    dn4, css4, pg4, ltl4, epnm4, an4, ld4, aan4, d4, ad4,
                    dn5, css5, pg5, ltl5, epnm5, an5, ld5, aan5, d5, ad5,
                    
                    
                    pool.blfs[1].number, pool.blfs[1].label,
                    pool.blfs[2].number, pool.blfs[2].label,
                    pool.blfs[3].number, pool.blfs[3].label,
                    pool.blfs[4].number, pool.blfs[4].label,
                    pool.blfs[5].number, pool.blfs[5].label,
                    pool.blfs[6].number, pool.blfs[6].label,
                    pool.blfs[7].number, pool.blfs[7].label,
                    pool.blfs[8].number, pool.blfs[8].label,
                    pool.blfs[9].number, pool.blfs[9].label,
                    pool.blfs[10].number, pool.blfs[10].label,
                    pool.blfs[11].number, pool.blfs[11].label,
                    pool.blfs[12].number, pool.blfs[12].label,
                    pool.blfs[13].number, pool.blfs[13].label,
                    pool.blfs[14].number, pool.blfs[14].label,
                    pool.blfs[15].number, pool.blfs[15].label,
                    pool.blfs[16].number, pool.blfs[16].label,
                    pool.blfs[17].number, pool.blfs[17].label,
                    pool.blfs[18].number, pool.blfs[18].label,
                    pool.blfs[19].number, pool.blfs[19].label,
                    pool.blfs[20].number, pool.blfs[20].label,
                    pool.blfs[21].number, pool.blfs[21].label,
                    pool.blfs[22].number, pool.blfs[22].label,
                    pool.blfs[23].number, pool.blfs[23].label,
                    pool.blfs[24].number, pool.blfs[24].label,
                    pool.blfs[25].number, pool.blfs[25].label,
                    pool.blfs[26].number, pool.blfs[26].label,
                    pool.blfs[27].number, pool.blfs[27].label,
                    pool.blfs[28].number, pool.blfs[28].label,
                    pool.blfs[29].number, pool.blfs[29].label,
                    pool.blfs[30].number, pool.blfs[30].label,

                    pool.speedDials[1].number, pool.speedDials[1].label,
                    pool.speedDials[2].number, pool.speedDials[2].label,
                    pool.speedDials[3].number, pool.speedDials[3].label,
                    pool.speedDials[4].number, pool.speedDials[4].label,
                    pool.speedDials[5].number, pool.speedDials[5].label,
                    pool.speedDials[6].number, pool.speedDials[6].label,
                    pool.speedDials[7].number, pool.speedDials[7].label,
                    pool.speedDials[8].number, pool.speedDials[8].label,
                    pool.speedDials[9].number, pool.speedDials[9].label,
                    pool.speedDials[10].number, pool.speedDials[10].label,
                    pool.speedDials[11].number, pool.speedDials[11].label,
                    pool.speedDials[12].number, pool.speedDials[12].label,
                    pool.speedDials[13].number, pool.speedDials[13].label,
                    pool.speedDials[14].number, pool.speedDials[14].label,
                    pool.speedDials[15].number, pool.speedDials[15].label,
                    pool.speedDials[16].number, pool.speedDials[16].label,
                    pool.speedDials[17].number, pool.speedDials[17].label,
                    pool.speedDials[18].number, pool.speedDials[18].label,
                    pool.speedDials[19].number, pool.speedDials[19].label,
                    pool.speedDials[20].number, pool.speedDials[20].label,
                    pool.speedDials[21].number, pool.speedDials[21].label,
                    pool.speedDials[22].number, pool.speedDials[22].label,
                    pool.speedDials[23].number, pool.speedDials[23].label,
                    pool.speedDials[24].number, pool.speedDials[24].label,
                    pool.speedDials[25].number, pool.speedDials[25].label,
                    pool.speedDials[26].number, pool.speedDials[26].label,
                    pool.speedDials[27].number, pool.speedDials[27].label,
                    pool.speedDials[28].number, pool.speedDials[28].label,
                    pool.speedDials[29].number, pool.speedDials[29].label,
                    pool.speedDials[30].number, pool.speedDials[30].label

                    );

            }
            ds.Tables.Add(export);


            //string AppLocation = "";
            
            saveFileDialog1.ValidateNames = true;
            saveFileDialog1.DereferenceLinks = false; // Will return .lnk in shortcuts.
            saveFileDialog1.Filter = "Excel |*.xlsx";
            saveFileDialog1.Title = "Export to Excel";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                string filepath = saveFileDialog1.FileName;

                // AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                // AppLocation = AppLocation.Replace("file:\\", "");
                // string date = DateTime.Now.ToShortDateString();
                // date = date.Replace("/", "_");
                // string filepath = AppLocation + "\\ExcelFiles\\" + "RECEIPTS_COMPARISON_" + date + ".xlsx";

                using (XLWorkbook wb = new XLWorkbook())
                {
                    for (int i = 0; i < ds.Tables.Count; i++)
                    {
                        wb.Worksheets.Add(ds.Tables[i], ds.Tables[i].TableName);
                    }
                    wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wb.Style.Font.Bold = true;
                    try
                    {
                        wb.SaveAs(filepath);
                    } 
                    catch (Exception ex)
                    { 
                    }
                }

            }
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if ((prefixBox.Text != "" && prefixBox.Text != "+1") && (comboBox1.Text != ""))
            {
                string[] site = comboBox1.Text.Split('-');
                parse.siteNumber = site[0];
                parse.siteName =  site[1];
                textBox.AppendText("SiteNumber:" + parse.siteNumber + Environment.NewLine);
                textBox.AppendText("SiteName:" + parse.siteName + Environment.NewLine);
                prefixBox.Enabled = false;
                comboBox1.Enabled = false;
                button4.Enabled = false;
                button1.Enabled = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
