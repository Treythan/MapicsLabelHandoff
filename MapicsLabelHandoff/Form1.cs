using Microsoft.VisualBasic.Devices;
using System.Text;

namespace MapicsLabelHandoff
{
    public partial class Form1 : Form
    {
        BarTender.Application btApp;
        BarTender.Format btFormat;
        Data data;
        string page;
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
        string serialNumber;
        string customer;
        public Form1(string[] args)
        {
            InitializeComponent();
            AddEventHandlers();
            InitPrint(args);
        }

        private void InitPrint(string[] args)
        {
            //args[0] = tabIndex
            StringBuilder sb = new StringBuilder();
            foreach (string arg in args) { sb.Append(arg); }
            string[] parsedArgs = sb.ToString().Split("||");

            switch (parsedArgs[0])
            {
                case "1":
                    tabMaster.SelectedIndex = 0;
                    ProcessProductReceipt(parsedArgs);
                    page = "1";
                    break;
                case "2":
                    tabMaster.SelectedIndex = 1;
                    ProcessSerializedReceipt(parsedArgs);
                    page = "2";
                    break;
            }
        }
        private void ProcessProductReceipt(string[] parsedArgs)
        {
            data = new Data().LoadData();
            //args format = pg,partnumber,partdesc,location,po,jn,uom
            partNumber = parsedArgs[1];
            partDescription = parsedArgs[2];
            location = parsedArgs[3];
            purchaseOrder = parsedArgs[4];
            jobNumber = parsedArgs[5];
            unitOfMeasure = parsedArgs[6];
            quantity = parsedArgs[7].Split(".")[0];
            labelAmount = "1";

            prodRecp1.Text = labelAmount;
            prodRecp2.Text = quantity;
            prodRecp6.Text = partNumber;
            prodRecp5.Text = partDescription;
            prodRecp4.Text = location;
            prodRecp9.Text = jobNumber;
            prodRecp8.Text = purchaseOrder;
            prodRecp7.Text = unitOfMeasure;
            prodRecp3.Text = data.lastReceiver;

            string lbl = "\\\\revginc.net\\AMB-GRP\\AMB-JNC_Orders\\Parts\\Labels\\4x4 Product Receipts Mapics.btw";

            btApp = GenBTIns();
            btApp.Visible = false;

            btFormat = btApp.Formats.Open(lbl, false, "");
        }

        private void ProcessSerializedReceipt(string[] parsedArgs)
        {
            data = new Data().LoadData();

            //args format = pg||partnumber||partdesc||location||jn||uom||customer||serial
            partNumber = parsedArgs[1];
            partDescription = parsedArgs[2];
            location = parsedArgs[3];
            jobNumber = parsedArgs[4];
            unitOfMeasure = parsedArgs[5];
            quantity = "1";
            customer = parsedArgs[6];
            serialNumber = parsedArgs[7];
            labelAmount = "1";

            sn1.Text = labelAmount;
            sn2.Text = quantity;
            sn3.Text = serialNumber;
            sn4.Text = customer;
            sn5.Text = data.lastReceiver;
            sn6.Text = partNumber;
            sn7.Text = partDescription;
            sn8.Text = location;
            sn9.Text = jobNumber;
            sn10.Text = data.lastPurchaseOrder;
            sn11.Text = unitOfMeasure;

            string lbl = "\\\\revginc.net\\AMB-GRP\\AMB-JNC_Orders\\Parts\\Labels\\4x4 Serialized Receipts Mapics.btw";

            btApp = GenBTIns();
            btApp.Visible = false;

            btFormat = btApp.Formats.Open(lbl, false, "");
        }

        private void FinishPrint()
        {
            switch (page)
            {
                case "1":
                    labelAmount = prodRecp1.Text;
                    quantity = prodRecp2.Text;
                    partNumber = prodRecp6.Text.ToUpper();
                    partDescription = prodRecp5.Text.ToUpper();
                    location = prodRecp4.Text.ToUpper();
                    purchaseOrder = prodRecp9.Text.ToUpper();
                    jobNumber = prodRecp8.Text.ToUpper();
                    unitOfMeasure = prodRecp7.Text.ToUpper();
                    receiver = prodRecp3.Text.ToUpper();

                    btFormat.SetNamedSubStringValue("PART_NUMBER", partNumber);
                    btFormat.SetNamedSubStringValue("PART_DESCRIPTION", partDescription);
                    btFormat.SetNamedSubStringValue("PURCHASE_ORDER", purchaseOrder);
                    btFormat.SetNamedSubStringValue("JOB_NUMBER", jobNumber);
                    btFormat.SetNamedSubStringValue("LOCATION", location);
                    btFormat.SetNamedSubStringValue("QUANTITY", quantity);
                    btFormat.SetNamedSubStringValue("RECEIVER", receiver);
                    btFormat.SetNamedSubStringValue("UOM", unitOfMeasure);
                    break;
                case "2":
                    labelAmount = sn1.Text;
                    quantity = sn2.Text;
                    partNumber = sn6.Text.ToUpper();
                    partDescription = sn7.Text.ToUpper();
                    location = sn8.Text.ToUpper();
                    purchaseOrder = sn10.Text.ToUpper();
                    jobNumber = sn9.Text.ToUpper();
                    unitOfMeasure = sn11.Text.ToUpper();
                    receiver = sn5.Text.ToUpper();
                    serialNumber = sn3.Text;
                    customer = sn4.Text.ToUpper();

                    btFormat.SetNamedSubStringValue("PART_NUMBER", partNumber);
                    btFormat.SetNamedSubStringValue("PART_DESCRIPTION", partDescription);
                    btFormat.SetNamedSubStringValue("PURCHASE_ORDER", purchaseOrder);
                    btFormat.SetNamedSubStringValue("JOB_NUMBER", jobNumber);
                    btFormat.SetNamedSubStringValue("LOCATION", location);
                    btFormat.SetNamedSubStringValue("QUANTITY", quantity);
                    btFormat.SetNamedSubStringValue("RECEIVER", receiver);
                    btFormat.SetNamedSubStringValue("UOM", unitOfMeasure);
                    btFormat.SetNamedSubStringValue("SERIAL", serialNumber);
                    btFormat.SetNamedSubStringValue("CUSTOMER", customer);
                    break;
            }
            for (int i = 0; i < int.Parse(labelAmount); i++)
            {
                btFormat.PrintOut(false, false);
            }

            data.lastReceiver = receiver;
            data.lastPurchaseOrder = purchaseOrder;
            data.SaveData();

            btFormat.Save();
            this.Close();
        }

        #region Helpers 

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

        private void AddEventHandlers()
        {
            foreach (Control tabPage in tabMaster.Controls)
            {
                foreach (Control control in tabPage.Controls)
                {
                    if (control is Button)
                    {
                        Button b = (Button)control;
                        b.Click += (s, e) =>
                        {
                            FinishPrint();
                        };
                    }

                    if (control is TextBox)
                    {
                        TextBox tb = (TextBox)control;

                        tb.KeyDown += (s, e) =>
                        {
                            if (e.KeyData == Keys.Enter)
                            {
                                FinishPrint();
                            }
                        };

                        tb.Click += (s, e) =>
                        {
                            tb.SelectAll();
                        };

                        tb.Enter += (s, e) =>
                        {
                            tb.SelectAll();
                        };
                    }
                }

            }
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                FinishPrint();
            }
        }
        #endregion
    }
}
