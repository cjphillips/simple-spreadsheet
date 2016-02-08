/* Colin Phillips
 * 11357836
 * CptS 322 - Assignment 10 (Assignment 10 - Handling Circular References)
 * October 29th, 2015
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetEngine;
using UndoRedoSystem;

namespace Spreadsheet_CPhillips
{
    public partial class Form1 : Form
    {
        public Spreadsheet mainSS;
        public SaveLoad MainSave;
        private int height = 50;
        private int width = 26;
        private int DefaultBGColor = -1;
        
        public Form1()
        {
            InitializeComponent();
            mainSS = new Spreadsheet(this.height, this.width, DefaultBGColor); // Creats a new spreadsheet of size 50 rows by 26 columns with the default window color (-1)
            MainSave = new SaveLoad();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            for (char c = 'A'; c != 'Z' + 1; c++) // Initializes column headers in datagrid
            {
                dataGridView1.Columns.Add(new DataGridViewTextBoxColumn() { HeaderText = c.ToString() } );
            }

            dataGridView1.Columns[0].Width = 100; // makes the first column header's width a little bigger than the default size

            for (int i = 1; i <= 50; i++) // Initializes row headers in datagrid
            {
                dataGridView1.Rows.Add(new DataGridViewRow());
                dataGridView1.Rows[i-1].HeaderCell.Value = i.ToString();
            }

            undoToolStripMenuItem.Enabled = false;
            undoToolStripMenuItem.Text = "Undo";
            redoToolStripMenuItem.Enabled = false;
            redoToolStripMenuItem.Text = "Redo";

            mainSS.CellPropertyChanged += Spreadsheet_PropertyChanged; // Subscribes the Grid's changing notification property to the Spreadsheets event
            MainSave.SavedPropertyChanged += Spreadsheet_PropertyChanged; // Subscribes the Grid's changing notification property to the Saved event
        }

        private void Spreadsheet_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Value") // continues when the cell's VALUE is being changed
            {
                dataGridView1[(sender as SpreadsheetCell).ColumnIndex, (sender as SpreadsheetCell).RowIndex].Value = (sender as SpreadsheetCell).Value;

                MainSave.Saved = false;
            }
            if (e.PropertyName == "BGColor") // continues here when the cell's back ground color is being changed
            {
                int newColor = (sender as SpreadsheetCell).BGColor;
                dataGridView1[(sender as SpreadsheetCell).ColumnIndex, (sender as SpreadsheetCell).RowIndex].Style.BackColor = System.Drawing.Color.FromArgb(newColor);

                MainSave.Saved = false;
            }
            if (e.PropertyName == "Saved" || e.PropertyName == "FileName")
            {
                if (MainSave.Saved)
                    SaveLabel.Text = "Saved - " + DateTime.Now; // Updates the saved label with the time/date it was saved
                else
                    SaveLabel.Text = "Not Saved."; 
            }

            undoToolStripMenuItem.Enabled = true; // Ensures that the undo button is enabled after changing a cell's property
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            dataGridView1[e.ColumnIndex, e.RowIndex].Value = mainSS[e.RowIndex, e.ColumnIndex].Text; // Sets converts the 'value' in the cell to the actual 'text'
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            SpreadsheetCell toEdit = mainSS[e.RowIndex, e.ColumnIndex];

            if (dataGridView1[e.ColumnIndex, e.RowIndex].Value != null)
            {
                string currentVal = toEdit.Value; // Stores the current value for later comparison
                string prev = toEdit.Text; // Stores the previous text

                string fromGrid = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
                mainSS[e.RowIndex, e.ColumnIndex].Text = fromGrid; // Updates the value of the current cell

                if (dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString() == prev) // If no change occured
                {
                    dataGridView1[e.ColumnIndex, e.RowIndex].Value = currentVal; // replaces the current cell with its old value
                    return;
                }

                mainSS.PushUndo(new CmdCollection(new RestoreText(mainSS[e.RowIndex, e.ColumnIndex], prev)));
                redoToolStripMenuItem.Enabled = false; // The redo stack is cleared after a new property is changed, so redo button should be disabled
                redoToolStripMenuItem.Text = "Redo";
            }
            else
            {
                toEdit.Text = "";
            }

            if (mainSS.UndoCount > 0)
                undoToolStripMenuItem.Text = "Undo - " + mainSS.PeekUndo; // Tells the UI which property will be restored
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit(); // Executes the Form1_FormClosing event (below)
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!MainSave.Saved) // Will Prompting the user to save if the current form's save is not up-to-date
            {
                DialogResult closing = MessageBox.Show("Save form before closing?", "Exit Spreadsheet", MessageBoxButtons.YesNoCancel);

                if (closing == DialogResult.Yes)
                {
                    saveToolStripMenuItem_Click(this, e); // Allows user to save the form
                    e.Cancel = false;
                }
                else if (closing == DialogResult.No) // User did not want to save the form
                    e.Cancel = false;
                else
                    e.Cancel = true;
            }
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainSS.PopUndo(); // Pops the top CmdCollection off the undo stack (also executes each command and restores the Spreadsheet)

            if (mainSS.UndoCount < 1) // disables the undo button if there are no more operations to undo
            {
                undoToolStripMenuItem.Enabled = false;
                undoToolStripMenuItem.Text = "Undo";
            }
            else
                undoToolStripMenuItem.Text = "Undo - " + mainSS.PeekUndo; // if still applicable, updates the property that will be restored

            redoToolStripMenuItem.Enabled = true;
            redoToolStripMenuItem.Text = "Redo - " + mainSS.PeekRedo; // also updates the redo operation that will be performed
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainSS.PopRedo(); // Pops the top CmdCollection off the redo stack (also executes each command and restores the Spreadsheet)

            if (mainSS.RedoCount < 1) // disables the redo button if there are no more operations to redo
            {
                redoToolStripMenuItem.Enabled = false;
                redoToolStripMenuItem.Text = "Redo";
            }
            else
                redoToolStripMenuItem.Text = "Redo - " + mainSS.PeekRedo; // if still applicable, updates the property that will be restored

            undoToolStripMenuItem.Text = "Undo - " + mainSS.PeekUndo; // also updates the undo operation that will be performed

            undoToolStripMenuItem.Enabled = true;
        }

        private void selectColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            color_changing(sender, true);
        }

        private void resetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            color_changing(sender, false);
        }

        private void greentoolStripMenuItem_Click(object sender, EventArgs e)
        {
            color_changing(sender, false);
        }

        private void orangeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            color_changing(sender, false);
        }

        private void redtoolStripMenuItem_Click(object sender, EventArgs e)
        {
            color_changing(sender, false);
        }

        private void greyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            color_changing(sender, false);
        }

        private void color_changing(object sender, bool custom)
        {
            ColorDialog SelectColor = new ColorDialog();
            if (custom) // user wants to select their own color -> prompts color dialogue
            {
                if (SelectColor.ShowDialog() != DialogResult.OK)
                    return;
            }
            else // the user chose one of the pre-selected colors
            {
                if (sender.ToString() == "Reset")
                    SelectColor.Color = System.Drawing.Color.FromArgb(-1); 
                if (sender.ToString() == "Green")
                    SelectColor.Color = System.Drawing.Color.FromArgb(192, 255, 192);
                else if (sender.ToString() == "Red")
                    SelectColor.Color = System.Drawing.Color.FromArgb(255, 128, 128);
                else if (sender.ToString() == "Orange")
                    SelectColor.Color = System.Drawing.Color.FromArgb(255, 192, 128);
                else if (sender.ToString() == "Grey")
                    SelectColor.Color = System.Drawing.Color.FromArgb(131, 130, 120);
            }

            DataGridViewSelectedCellCollection CellsChanging = dataGridView1.SelectedCells; // Gets the collection of selected cells
            CmdCollection BGRestore = new CmdCollection();

            if (CellsChanging != null) // Ensures that there is at least one selected cell
            {
                foreach (DataGridViewCell cell in CellsChanging) // Foreach cell that has been selected, notify the spreadsheet of that cells property change
                {
                    SpreadsheetCell Changing = mainSS[cell.RowIndex, cell.ColumnIndex];
                    RestoreBGColor newRestore = new RestoreBGColor(Changing, Changing.BGColor);

                    Changing.BGColor = SelectColor.Color.ToArgb(); // Converts the chosen color to ARBG integer format
                    BGRestore.Add(newRestore); 
                }

                mainSS.PushUndo(BGRestore); // Adds the new restore collection
                undoToolStripMenuItem.Text = "Undo - " + mainSS.PeekUndo; // Tells the UI which property will be undone
                redoToolStripMenuItem.Enabled = false;
                redoToolStripMenuItem.Text = "Redo";
            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog SaveAs = new SaveFileDialog();
            SaveAs.Filter = "*.txt | *.txt"; // Only allows .xml formats to save to
            SaveAs.Title = "Save Spreadsheet";

            if (SaveAs.ShowDialog() == DialogResult.OK) // Prompts the user to enter a new file name
            {
                try
                {
                    MainSave.Save(SaveAs.FileName, mainSS); // attempts to save the file
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Error - Could not save to File", MessageBoxButtons.OK);
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MainSave.FileName == "") 
                saveAsToolStripMenuItem_Click(sender, e); // redirects to the Save As menu if no file has been chosen yet
            else
            {
                MainSave.Save(MainSave.FileName, mainSS); // Saves the form with the current file
            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult Ensure = MessageBox.Show("Continue?", "New Spreadsheet", MessageBoxButtons.YesNo);

            if (Ensure == DialogResult.Yes) // Ensures the user meant to start a new form
            {
                if (!MainSave.Saved) // If the current form's state has not been saved, will ask if it should be 
                {
                    DialogResult closing = MessageBox.Show("Save Spreadsheet?", "New Spreadsheet", MessageBoxButtons.YesNoCancel);

                    if (closing == DialogResult.Yes)
                        saveToolStripMenuItem_Click(this, e); // Executes saving of form
                    else if (closing == DialogResult.Cancel)
                        return; // cancels the new form execution
                }

                mainSS = new Spreadsheet(this.height, this.width, DefaultBGColor); // Reinitializes the spreadsheet
                dataGridView1.Rows.Clear(); // clears all rows in the UI
                dataGridView1.Columns.Clear(); // clears all columns in the UI
                Form1_Load(sender, e); // Reinitializes the DataGridView
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!MainSave.Saved) // If the current session has not been saved, will prompt to save before loading
            {
                DialogResult loading = MessageBox.Show("Save before loading?", "Load Spreadsheet", MessageBoxButtons.YesNoCancel);

                if (loading == DialogResult.Yes)
                    saveAsToolStripMenuItem_Click(this, e); // saves the current form
                else if (loading == DialogResult.Cancel)
                    return;
            }

            OpenFileDialog OpenFrom = new OpenFileDialog(); // prompts the user to enter a file to open
            if (OpenFrom.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    mainSS = new Spreadsheet(this.height, this.width, DefaultBGColor); // Reinitializes the spreadsheet
                    dataGridView1.Rows.Clear(); // clears all rows in the UI
                    dataGridView1.Columns.Clear(); // clears all columns in the UI
                    Form1_Load(sender, e); // Reinitializes the DataGridView

                    MainSave.Load(OpenFrom.FileName, mainSS); // Loads the info from the input file
                    MainSave.Saved = true; // Form comes in as saved
                    SaveLabel.Text = "Load Successful.";
                    undoToolStripMenuItem.Enabled = false;
                    undoToolStripMenuItem.Text = "Undo"; // loading resets the undo system
                    redoToolStripMenuItem.Enabled = false;
                    redoToolStripMenuItem.Text = "Redo"; // loading resets the redo system
                }
                catch (Exception err) // catches an exception if the file failed to load
                {
                    DialogResult CannotOpen = MessageBox.Show(err.Message, "Error - Cannot Open File", MessageBoxButtons.OK);
                }
            }
        }
    }
}
