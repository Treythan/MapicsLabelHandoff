using System.Text;

namespace MapicsLabelHandoff
{
    public partial class Form1 : Form
    {
        BarTender.Application btApp;
        BarTender.Format btFormat;
        string[] parsedArgs;
        string partNumber;
        string partDescription;
        string location;
        string purchaseOrder;
        string jobNumber;
        string unitOfMeasure;
        string quantity;
        string receiver;
        string labelAmount;
        string lastReceiverPath = AppContext.BaseDirectory + "lastReceiver.txt";
        public Form1(string[] args)
        {
            InitializeComponent();
            InitPrint(args);
        }

        private void InitPrint(string[] args)
        {
            //args format = partnumber,partdesc,location,po,jn,uom
            StringBuilder sb = new StringBuilder();
            foreach (string arg in args) { sb.Append(arg); }
            string[] parsedArgs = sb.ToString().Split("||");
            partNumber = parsedArgs[0];
            partDescription = parsedArgs[1];
            location = parsedArgs[2];
            purchaseOrder = parsedArgs[3];
            jobNumber = parsedArgs[4];
            unitOfMeasure = parsedArgs[5];
            quantity = parsedArgs[6].Split(".")[0];
            labelAmount = "1";

            textBox1.Text = labelAmount;
            textBox2.Text = quantity;
            textBox6.Text = partNumber;
            textBox5.Text = partDescription;
            textBox4.Text = location;
            textBox9.Text = purchaseOrder;
            textBox8.Text = jobNumber;
            textBox7.Text = unitOfMeasure;
            textBox3.Text = File.ReadAllText(lastReceiverPath);

            string lbl = "\\\\revginc.net\\AMB-GRP\\AMB-JNC_Orders\\Parts\\Labels\\4x4 Product Receipts Mapics.btw";

            btApp = GenBTIns();
            btApp.Visible = false;

            btFormat = btApp.Formats.Open(lbl, false, "");
        }

        private BarTender.Application GenBTIns()
        {
            try
            {
                return Marshall.GetActiveObject("BarTender.Application") as BarTender.Application;
            }
            catch (Exception e)
            {
                return new BarTender.Application();
            }
        }

        private void FinishPrint()
        {
            labelAmount = textBox1.Text;
            quantity = textBox2.Text;
            partNumber = textBox6.Text.ToUpper();
            partDescription = textBox5.Text.ToUpper();
            location = textBox4.Text.ToUpper();
            purchaseOrder = textBox9.Text.ToUpper();
            jobNumber = textBox8.Text.ToUpper();
            unitOfMeasure = textBox7.Text.ToUpper();
            receiver = textBox3.Text == "" ? File.ReadAllText(lastReceiverPath) : textBox3.Text.ToUpper();

            btFormat.SetNamedSubStringValue("PART_NUMBER", partNumber);
            btFormat.SetNamedSubStringValue("PART_DESCRIPTION", partDescription);
            btFormat.SetNamedSubStringValue("PURCHASE_ORDER", purchaseOrder);
            btFormat.SetNamedSubStringValue("JOB_NUMBER", jobNumber);
            btFormat.SetNamedSubStringValue("LOCATION", location);
            btFormat.SetNamedSubStringValue("QUANTITY", quantity);
            btFormat.SetNamedSubStringValue("RECEIVER", receiver);
            btFormat.SetNamedSubStringValue("UOM", unitOfMeasure);

            for (int i = 0; i < int.Parse(labelAmount); i++)
            {
                btFormat.PrintOut(false, false);
            }

            File.WriteAllText(lastReceiverPath, receiver);

            btFormat.Save();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FinishPrint();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                FinishPrint();
            }
        }
        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                FinishPrint();
            }
        }
        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                FinishPrint();
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.SelectAll();
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            textBox1.SelectAll();
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.SelectAll();
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            textBox2.SelectAll();
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            textBox3.SelectAll();
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            textBox3.SelectAll();
        }
    }
}
