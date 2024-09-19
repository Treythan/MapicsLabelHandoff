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
        public Form1(string[] args)
        {
            InitializeComponent();
            AddEventHandlers();
            ProcessProductReceipt(args);
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

            _labels.Text = labelAmount;
            _quantity.Text = quantity;
            _partNumber.Text = partNumber;
            _partDescription.Text = partDescription;
            _location.Text = location;
            _jobNumber.Text = jobNumber;
            _purchaseOrder.Text = purchaseOrder;
            _unitOfMeasure.Text = unitOfMeasure;
            _initials.Text = data.lastReceiver;

            string lbl = "\\\\revginc.net\\AMB-GRP\\AMB-JNC_Orders\\Parts\\Labels\\4x4 Product Receipts Mapics.btw";

            btApp = GenBTIns();
            btApp.Visible = false;

            btFormat = btApp.Formats.Open(lbl, false, "");
        }

        private void FinishPrint()
        {
            labelAmount = _labels.Text;
            quantity = _quantity.Text;
            partNumber = _partNumber.Text.ToUpper();
            partDescription = _partDescription.Text.ToUpper();
            location = _location.Text.ToUpper();
            purchaseOrder = _purchaseOrder.Text.ToUpper();
            jobNumber = _jobNumber.Text.ToUpper();
            unitOfMeasure = _unitOfMeasure.Text.ToUpper();
            receiver = _initials.Text.ToUpper();

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
            foreach (Control c in this.Controls)
            {
                if (c is Button)
                {
                    Button b = (Button)c;
                    b.Click += (s, e) =>
                    {
                        FinishPrint();
                    };
                }

                if (c is TextBox)
                {
                    TextBox tb = (TextBox)c;

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
