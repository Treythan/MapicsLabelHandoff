namespace MapicsLabelHandoff
{
    public partial class Form2 : Form
    {
        BarTender.Application btApp;
        BarTender.Format btFormat;
        Data data;
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
        public Form2(string[] args)
        {
            InitializeComponent();
            AddEventHandlers();
            ProcessSerializedProductReceipt(args);
        }
        private void ProcessSerializedProductReceipt(string[] parsedArgs)
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

            _labels.Text = labelAmount;
            _quantity.Text = quantity;
            _serial.Text = serialNumber;
            _customer.Text = customer;
            _initials.Text = data.lastReceiver;
            _partNumber.Text = partNumber;
            _partDescription.Text = partDescription;
            _location.Text = location;
            _jobNumber.Text = jobNumber;
            _purchaseOrder.Text = data.lastPurchaseOrder;
            _unitOfMeasure.Text = unitOfMeasure;

            string lbl = "\\\\revginc.net\\AMB-GRP\\AMB-JNC_Orders\\Parts\\Labels\\4x4 Serialized Receipts Mapics.btw";

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
            serialNumber = _serial.Text;
            customer = _customer.Text.ToUpper();

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
