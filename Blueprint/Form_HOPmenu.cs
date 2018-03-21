using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; //if this is not working, you need to add a reference to the excel interop

namespace Blueprint
{
    public partial class Form_HOPmenu : Form
    {
        private Project project;

        //excel variables
        private Excel.Application oXL;
        private Excel._Workbook oWB;

        public Form_HOPmenu()
        {
            InitializeComponent();
            project = new Project();
            populate_comboboxes();
            initial_form_state();
        }


        private void button_done_Click(object sender, EventArgs e)
        {
            update_project_data();

            //start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = true;

            //get a new workbook.
            oWB = oXL.Workbooks.Add();

            Traveler traveler = new Traveler(project, oWB);
            traveler.generate();

            MaterialsID materials = new MaterialsID(project, oWB);
            materials.generate();

            WelderID welders = new WelderID(project, oWB);
            welders.generate();
        }


        private void initial_form_state()
        {
            this.groupBox_vesselDesign.Enabled = false;

            this.radioButton_horizontal.Checked = true;
            this.textBox_vesselLength.Text = "36";
            this.radioButton_OAL.Checked = true;
            this.radioButton_cs.Checked = true;
            this.textBox_tMinShell.Text = "0.117";
            this.textBox_tMinHeads.Text = "0.117";
            this.textBox_tMinAfterForming.Text = "0.2818";
            this.checkBox_fv.Checked = true;
            this.checkBox_testWithWater.Checked = true;
            this.checkBox_n2Charge.Checked = true;
            this.textBox_paintSpecification.Text = "SANDBLAST SSPC-SP6, RVS HYPER BLUE EPOXY";

            this.textBox_date.Enabled = false;
            this.textBox_date.Text = DateTime.Now.ToString("M/d/yy");
        }


        private void update_project_data()
        {
            project.projectNumber = this.textBox_projectNumber.Text;
            project.serialNumber = this.textBox_serialNumber.Text;
            project.evapcoNumber = this.textBox_evapcoNumber.Text;
            project.drawingNumber = this.textBox_drawingNumber.Text;
            project.revisionNumber = this.textBox_revisionNumber.Text;
            project.approvalStatus = this.comboBox_approvalStatus.Text;
            project.customer = this.textBox_customer.Text;
            project.poNumber = this.textBox_poNumber.Text;

            project.vessel.isHorizontal = this.radioButton_horizontal.Checked;
            project.vessel.outerDiameter = Double.Parse(this.comboBox_vesselOD.Text);
            project.vessel.lengthSmSm = Double.Parse(this.textBox_vesselLength.Text);
            project.vessel.isOAL = this.radioButton_OAL.Checked;
            project.vesselType = this.comboBox_vesseType.Text;
            project.vessel.isSS = this.radioButton_ss.Checked;
            project.designPressure = this.comboBox_pressure.Text;
            project.maxTemperature = this.comboBox_designTemp.Text;
            project.minTemperature = this.comboBox_mdmt.Text;
            project.asmeEdition = this.comboBox_asmeEdition.Text;
            project.rtLong = this.comboBox_rtLong.Text;
            project.rtGirth = this.comboBox_rtGirth.Text;
            project.testPressure = this.comboBox_testPressure.Text;
            project.caShell = this.comboBox_caShell.Text;
            project.caHeads = this.comboBox_caHeads.Text;
            project.tMinShell = this.textBox_tMinShell.Text;
            project.tMinHeads = this.textBox_tMinHeads.Text;
            project.tMinAfterForming = this.textBox_tMinAfterForming.Text;
            project.paintSpecification = this.textBox_paintSpecification.Text;
            project.getsPWHT = this.checkBox_pwht.Checked;
            project.getsFullVacuum = this.checkBox_fv.Checked;
            project.getsHydrotest = this.checkBox_testWithWater.Checked;
            project.getsN2charge = this.checkBox_n2Charge.Checked;
            project.getsTestClosures = this.checkBox_testClosures.Checked;

            project.initials = this.textBox_initials.Text;
            project.date = this.textBox_date.Text;
        }


        private void populate_comboboxes()
        {
            populate_approval_status();
            populate_hop_style();
            populate_hop_size();
            populate_hop_options();
            populate_vessel_od();
            populate_vessel_type();
            populate_design_pressure();
            populate_design_temperature();
            populate_mdmt();
            populate_asme_edition();
            populate_rt_long();
            populate_rt_girth();
            populate_test_pressure();
            populate_ca_shell();
            populate_ca_heads();
        }


        private void populate_approval_status()
        {
            ComboBox box = this.comboBox_approvalStatus;
            box.Items.Clear();

            box.Items.Add("FOR APPROVAL");
            box.Items.Add("CERTIFIED");

            box.SelectedIndex = 0;
        }


        private void populate_hop_style()
        {
            ComboBox box = this.comboBox_hopStyle;
            box.Items.Clear();

            box.Items.Add("HOP");
            box.Items.Add("MPC-LT-OP");
            box.Items.Add("MPC-RT-OP");
            box.Items.Add("MRP-H-OP");
            box.Items.Add("MRP-V-OP");

            box.SelectedIndex = 0;
        }


        private void populate_hop_size()
        {
            ComboBox box = this.comboBox_hopSize;
            box.Items.Clear();

            box.Items.Add("8-36-400");
            box.Items.Add("10-36-400");

            box.SelectedIndex = 0;
        }


        private void populate_hop_options()
        {
            ComboBox box = this.comboBox_hopOptions;
            box.Items.Clear();

            box.Items.Add("");
            box.Items.Add("H");
            box.Items.Add("RT");
            box.Items.Add("RT-H");
            box.Items.Add("SS");
            box.Items.Add("SS-H");
            box.Items.Add("SS-RT");
            box.Items.Add("SS-RT-H");

            box.SelectedIndex = 0;
        }


        private void populate_vessel_od()
        {
            ComboBox box = this.comboBox_vesselOD;
            box.Items.Clear();

            box.Items.Add("");
            box.Items.Add("8.625");
            box.Items.Add("10.75");

            box.SelectedIndex = 1;
        }


        private void populate_vessel_type()
        {
            ComboBox box = this.comboBox_vesseType;
            box.Items.Clear();

            box.Items.Add("");
            box.Items.Add("OIL POT");

            box.SelectedIndex = 1;
        }


        private void populate_design_pressure()
        {
            ComboBox box = this.comboBox_pressure;
            box.Items.Clear();

            box.Items.Add("");
            box.Items.Add("250");
            box.Items.Add("300");
            box.Items.Add("400");

            box.SelectedIndex = 3;
        }


        private void populate_design_temperature()
        {
            ComboBox box = this.comboBox_designTemp;
            box.Items.Clear();

            box.Items.Add("");
            box.Items.Add("300");

            box.SelectedIndex = 1;
        }


        private void populate_mdmt()
        {
            ComboBox box = this.comboBox_mdmt;
            box.Items.Clear();

            box.Items.Add("");
            box.Items.Add("-50");

            box.SelectedIndex = 1;
        }


        private void populate_asme_edition()
        {
            ComboBox box = this.comboBox_asmeEdition;
            box.Items.Clear();

            box.Items.Add("2017");

            box.SelectedIndex = 0;
        }


        private void populate_rt_long()
        {
            ComboBox box = this.comboBox_rtLong;
            box.Items.Clear();

            box.Items.Add("NONE");
            box.Items.Add("SPOT");
            box.Items.Add("FULL");

            box.SelectedIndex = 0;
        }


        private void populate_rt_girth()
        {
            ComboBox box = this.comboBox_rtGirth;
            box.Items.Clear();

            box.Items.Add("NONE");
            box.Items.Add("SPOT");
            box.Items.Add("FULL");

            box.SelectedIndex = 0;
        }


        private void populate_test_pressure()
        {
            ComboBox box = this.comboBox_testPressure;
            box.Items.Clear();

            box.Items.Add("");
            box.Items.Add("520");

            box.SelectedIndex = 1;
        }


        private void populate_ca_shell()
        {
            ComboBox box = this.comboBox_caShell;
            box.Items.Clear();

            box.Items.Add("0");

            box.SelectedIndex = 0;
        }


        private void populate_ca_heads()
        {
            ComboBox box = this.comboBox_caHeads;
            box.Items.Clear();

            box.Items.Add("0");

            box.SelectedIndex = 0;
        }
    }
}
