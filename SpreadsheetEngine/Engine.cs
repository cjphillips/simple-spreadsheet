/* Colin Phillips
 * 11357836
 * CptS 322 - Assignment 10 (Assignment 10 - Handling Circular References)
 * October 29th, 2015
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using System.ComponentModel;
using ExpressionTree;
using UndoRedoSystem;
using System.IO;
using System.Xml;

namespace SpreadsheetEngine
{
    public abstract class SpreadsheetCell : INotifyPropertyChanged
    {
        private int rowIndex;
        private int columnIndex;
        private string textStr;
        protected string valueStr; // This value is protected -> only accessible through inheritance
        protected int BackColor; // Allows inheriting classes to set the default cell color
        private HashSet<string> varList;

        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        public SpreadsheetCell(int setRowIndex, int setColumnIndex)
        {
            rowIndex = setRowIndex; // permanently sets the row and column indicies
            columnIndex = setColumnIndex;
            textStr = "";
            valueStr = "";
            BackColor = 0;
            varList = null;
        }

        public int RowIndex { get { return rowIndex; } }
        public int ColumnIndex { get { return columnIndex; } }

        
        public string Text
        {
            get { return textStr; }

            set
            {
                if (value != textStr) // ensures that the string passed is different than the current text
                {
                    textStr = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("Value")); // fires the value property changed event
                }
            }
        }
        public string Value { get { return valueStr; } }

        public int BGColor
        {
            get { return BackColor; }

            set
            {
                if (value != BackColor) // ensures that the color is actually changing
                {
                    BackColor = value;
                    PropertyChanged(this, new PropertyChangedEventArgs("BGColor")); // Fires the spreadsheet's property changed (for BG Color)
                }
            }
        }

        internal HashSet<string> VarList // Only the DLL should be able to access a cell's variable list -> otherwise it's private
        {
            get { return this.varList; }
            set { this.varList = value; }
        }

        public bool IsDefault()
        {
            if (this.BackColor == -1) // default color (window: -1)
            {
                if (this.Text == "") // default text (empty string)
                    return true;
            }

            return false;
        }
    }

    class Cell : SpreadsheetCell
    {
        public Cell(int setRowIndex, int setColIndex, int DefaultCellColor) : base(setRowIndex, setColIndex) 
        {
            this.BackColor = DefaultCellColor;
        }

        public Cell(SpreadsheetCell from, string newVal) : base(from.RowIndex, from.ColumnIndex)
        {
            this.valueStr = newVal;
            this.Text = from.Text;
            this.BackColor = from.BGColor;
            this.VarList = from.VarList;
        }

        public Cell(SpreadsheetCell from) : base(from.RowIndex, from.ColumnIndex)
        {
            this.valueStr = from.Value;
            this.Text = from.Text;
            this.BackColor = from.BGColor;
            this.VarList = from.VarList;
        }
    }

    public class Spreadsheet
    {
        enum ErrType
        {
            None, SelfRef, CircRef, BadRef, DivZero, InvExp
        }

        private SpreadsheetCell[,] cellArr; // 2D array of Cells
        Dictionary<string, HashSet<string>> refTable = new Dictionary<string, HashSet<string>>();
        private int height; // Number of rows
        private int width; // Number of columns
        private int StartCellColor; // Stores the default passed color (for BG Color information)
        public event PropertyChangedEventHandler CellPropertyChanged = delegate { };
        private Stack<CmdCollection> UndoStack; // Will store all undo commands
        private Stack<CmdCollection> RedoStack; // Will store all redo commands

        public Spreadsheet(int numRows, int numColumns, int DefaultCellColor)
        {
            cellArr = new Cell[numRows, numColumns];
            width = numColumns;
            height = numRows;

            for (int i = 0; i < numRows; i++) // Initializes Rows
            {
                for (int j = 0; j < numColumns; j++) // Initializes Columns
                {
                    cellArr[i, j] = new Cell(i, j, DefaultCellColor); // Creates a new cell with it specific row and column indicies
                    cellArr[i, j].PropertyChanged += detect_PropertyChanged; // Subscribes the detect_PropertyChanged to each cell contained
                }
            }

            UndoStack = new Stack<CmdCollection>();
            RedoStack = new Stack<CmdCollection>();

            StartCellColor = DefaultCellColor; // Tells the cells which color to initially be
        }

        public SpreadsheetCell this[int row, int column] // Indexing Overload
        {
            get
            {
                if (row > height || column > width) // returns null if the passed values are outside of the height and/or width of the spreadsheet
                    return null;
                else
                    return cellArr[row, column]; // returns the cell
            }
        }

        public SpreadsheetCell GetCell(int row, int column)
        {
            if (row > height || column > width) // returns null if the passed values are outside of the height and/or width of the spreadsheet
                return null;
            else
                return cellArr[row, column]; // returns the cell
        }

        public int RowCount { get { return height; } } // returns the number of rows
        public int ColumnCount { get { return width; } } // returns the number of columns
        public int DefaultColor { get { return StartCellColor; } }

        private void detect_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            SpreadsheetCell toAlter = (sender as SpreadsheetCell);

            if (e.PropertyName == "Value")
            {
                UpdateCellValue(ref toAlter); // Updates the Cell and all of the cells that reference it
            }
            if (e.PropertyName == "BGColor")
            {
                CellPropertyChanged(toAlter, new PropertyChangedEventArgs("BGColor")); // notifies the UI that the spreadsheet is changing a cells background color
            }
        }

        private void UpdateCellValue(ref SpreadsheetCell cell)
        {
            var mainTree = new ExpTree(); // Initializes a new expression tree to build the cell's expression
            ErrType Error = ErrType.None;

            if (cell.Text != "" && cell.Text[0] != '=') // not an expression, simply a text value
            {
                cellArr[cell.RowIndex, cell.ColumnIndex] = new Cell(cell, cell.Text);
                add_remove(cell, mainTree, true);
            }
            else
            {
                if (cell.Text != "")
                    mainTree.Expression = cell.Text.Substring(1).Replace(" ", ""); // Build the expression tree with the cell's text (minus the '=') :: Also ignores whitespaces
                else
                    mainTree.Expression = cell.Text;

                add_remove(cell, mainTree, true); // Removes all variables cooresponding to the old tree

                cell.VarList = GetVarNames(cell.Text); // Will Return all found variables in the new expression

                Error = add_remove(cell, mainTree, false);

                if (Error != ErrType.None) // Notifies the UI that there is an error in one of the cells that the expression references
                {
                    return; // Exits the function before executing anything else, error display has already been taken care of at this point
                }

                try
                {
                    cellArr[cell.RowIndex, cell.ColumnIndex] = new Cell(cell, mainTree.Eval().ToString()); // Attempts to evaluate the expression, placing it into a new cell
                }
                catch (DivideByZeroException) // User tried to divide by zero
                {
                    CheckErr(cell, ErrType.DivZero, CellToString(cell));
                    UpdateErrorReferecedBy(cell, ErrType.DivZero, CellToString(cell));
                    return;
                }
                catch (NullReferenceException) // Input not regonized / invalid expression
                {
                    if (cell.Text == "")
                    {
                        cellArr[cell.RowIndex, cell.ColumnIndex] = new Cell(cell, ""); // if the cell was deleted or reset, this will set the cell to an empty value (caught by expression tree as null)
                    }
                    else
                    {
                        CheckErr(cell, ErrType.InvExp, CellToString(cell));
                        UpdateErrorReferecedBy(cell, ErrType.InvExp, CellToString(cell)); // Notifies UI that an invalid expression has been entered
                        return;
                    }
                }
            }

            cellArr[cell.RowIndex, cell.ColumnIndex].PropertyChanged += detect_PropertyChanged; // Reassigns the the detect_property function to the cell's delegate

            CellPropertyChanged(cellArr[cell.RowIndex, cell.ColumnIndex], new PropertyChangedEventArgs("Value")); // fires the event that notifies the GUI of a change

            UpdateReferencedBy(cell); // Updates all cells that reference this cell
        }

        private void UpdateErrorReferecedBy(SpreadsheetCell cell, ErrType Check, string root)
        {
            /* This function will update all of the cells that reference the passed cell IFF an error has been detected 
             * Cells that contain the error and cells that reference cells containing an error will use this function */
            if (this.refTable.ContainsKey(CellToString(cell))) // Ensures that the current cell has other cells that reference it
            {
                for (int i = 0; i < this.refTable[CellToString(cell)].Count; i++)
                {
                    SpreadsheetCell nextCell = StringToCell(this.refTable[CellToString(cell)].ElementAt<string>(i)); // Looks up the next cell that references the current cell
                    if (Check == ErrType.SelfRef && CellToString(nextCell) == root) // Stop self-referencing loops
                        break; // Updates all cells that reference it
                    else if (Check == ErrType.CircRef && CellToString(nextCell) == root) // stops circular-referencing loops
                        break;
                    else
                    {
                        CheckErr(nextCell, Check, root);
                        UpdateErrorReferecedBy(nextCell, Check, root); // Continues updated all cells that reference THIS cell
                    }
                }
            }
        }

        private void UpdateReferencedBy(SpreadsheetCell cell)
        {
            /*This function will update all of the cells that reference the passed cell*/
            if (this.refTable.ContainsKey(CellToString(cell))) // Ensures that the current cell has other cells that reference it
            {
                for (int i = 0; i < this.refTable[CellToString(cell)].Count; i++)
                {
                    SpreadsheetCell nextCell = StringToCell(this.refTable[CellToString(cell)].ElementAt<string>(i)); // Looks up the next cell that references the current cell
                    UpdateCellValue(ref nextCell);
                }
            }
        }

        private ErrType add_remove(SpreadsheetCell toAlter, ExpTree mainTree, bool removing)
        {
            /*Adding to and removing from the reference table occurs in this function*/
            ErrType Error = ErrType.None;

            if (toAlter.VarList != null && toAlter.VarList.Count > 0)
            {
                string referencedBy = CellToString(toAlter);
                if (removing)
                {
                    foreach (string referencedCell in toAlter.VarList) // Removes all variables from the old tree
                    {
                        if (refTable.ContainsKey(referencedCell))
                        {
                            if (refTable[referencedCell].Contains(referencedBy))
                                refTable[referencedCell].Remove(referencedBy); // Removes the current cell from any other cells referencing hash
                            if (refTable[referencedCell].Count < 1) // If an entry in the table has no cells referencing it, then it is removed
                                refTable.Remove(referencedCell);
                        }
                    }

                    toAlter.VarList.Clear(); // Empty that variable list (will be rebuild below)
                }
                else // Adding value to the reference this
                {
                    foreach (string s in toAlter.VarList) // Updates the reference table with all of the referenced cells (variables in the expTree's context)
                    {
                        double CellValue = 0.0;
                        SpreadsheetCell next = StringToCell(s);
                        if (next != null)
                        {
                            if (s == CellToString(toAlter)) // SELF-REFERENCING
                            {
                                Error = ErrType.SelfRef;
                                CheckErr(toAlter, Error, CellToString(toAlter));
                                UpdateErrorReferecedBy(toAlter, ErrType.SelfRef, CellToString(toAlter)); // Updates all cells referencing this cell that there is a value
                                return ErrType.SelfRef;
                            }
                            else if (next.Value.Contains("REF")) // Won't check for already occuring errors (in referenced cell) 
                            {
                                if (next.Value.Contains("=<BAD_REF")) // If this cell REFERENCES a cell that contains a bad_ref error
                                {
                                    CheckErr(toAlter, ErrType.BadRef, s);
                                    UpdateErrorReferecedBy(toAlter, ErrType.BadRef, s);
                                    Error = ErrType.BadRef;
                                }
                                else if (next.Value.Contains("=<SELF_REF")) // If this cell REFERENCES a cell that contains a self_ref error
                                {
                                    CheckErr(toAlter, ErrType.SelfRef, s);
                                    UpdateErrorReferecedBy(toAlter, ErrType.SelfRef, CellToString(toAlter));
                                    Error = ErrType.SelfRef;
                                }
                                else if (next.Value.Contains("=<CIRC_REF"))
                                {
                                    CheckErr(toAlter, ErrType.CircRef, s);
                                    UpdateErrorReferecedBy(toAlter, ErrType.CircRef, CellToString(toAlter));
                                    Error = ErrType.CircRef;
                                }
                            }
                            if (next.Text != "")
                            {
                                Double.TryParse(next.Value, out CellValue); // Gets the cell's value
                                mainTree.SetVar(s, CellValue); // Sets the variable in the expression tree's dictionary (0 if not yet set)
                            }
                            if (refTable.ContainsKey(s)) // If The variable already has references, just add to its hash
                                refTable[s].Add(referencedBy);
                            else // Otherwise create the new variable key with a new list containing the cell that references it
                                refTable.Add(s, new HashSet<string>() { referencedBy });
                        }
                        else // If Cell parsing return null (cell not recovered), the there is a bad reference
                        {
                            Error = ErrType.BadRef;
                            CheckErr(toAlter, Error, CellToString(toAlter));
                            UpdateErrorReferecedBy(toAlter, ErrType.BadRef, CellToString(toAlter));
                            return ErrType.BadRef;
                        }
                    }
                    if (Error == ErrType.CircRef)
                        return Error;

                    if (Error != ErrType.SelfRef && CheckCircularRef(toAlter, CellToString(toAlter))) // Checks for circular references here ***
                    {
                        Error = ErrType.CircRef;
                        CheckErr(toAlter, Error, CellToString(toAlter));
                        UpdateErrorReferecedBy(toAlter, ErrType.CircRef, CellToString(toAlter));
                    }
                }
            }

            return Error;
        }

        private void CheckErr(SpreadsheetCell cell, ErrType err, string root)
        {
            /* This function will check the passed error and update the cell accordingly
             * This function will also update cells referencing an error with the location
               where the error is occuring (exluding Circular references)
             * Reinitializes cell delegate */
               
            if (cell.Text == "") // Skips cells that have been reset
                return;
            if (err == ErrType.BadRef) // BAD REFERENCES
            {
                if (CellToString(cell) == root)
                    cell = new Cell(cell, "=<BAD_REF::SRC>"); // Plants the source of the error
                else
                {
                    cell = new Cell(cell, "=<BAD_REF::AT[" + root + "]"); // Updates the cell's value with the location of the error
                }
            }
            else if (err == ErrType.SelfRef) // SELF REFERENCES
            {
                if (CellToString(cell) == root)
                    cell = new Cell(cell, "=<SELF_REF::SRC>");
                else
                {
                    cell = new Cell(cell, "=<SELF_REF::AT[" + root + "]");
                }
            }
            else if (err == ErrType.CircRef) // CIRCULAR REFERENCES
            {
                cell = new Cell(cell, "=<CIRC_REF>");
            }
            else if (err == ErrType.DivZero) // DIVISION BY ZERO
            {
                if (CellToString(cell) == root)
                    cell = new Cell(cell, "=<DIV_ZERO::SRC>");
                else
                {
                    cell = new Cell(cell, "=<DIV_ZERO::AT[" + root + "]");
                }
            }
            else if (err == ErrType.InvExp) // INVALID EXPRESSIONS
            {
                if (CellToString(cell) == root)
                    cell = new Cell(cell, "=<INV_EXP::SRC>");
                else
                {
                    cell = new Cell(cell, "=<INV::EXP::AT[" + root + "]");
                }
            }

            if (err != ErrType.None)
            {
                cellArr[cell.RowIndex, cell.ColumnIndex] = new Cell(cell, cell.Value);
                cellArr[cell.RowIndex, cell.ColumnIndex].PropertyChanged += detect_PropertyChanged; // Reassigns the the detect_property function to the cell's delegate

                CellPropertyChanged(cellArr[cell.RowIndex, cell.ColumnIndex], new PropertyChangedEventArgs("Value")); // fires the event that notifies the GUI of a change
            }
        }
        
        private bool CheckCircularRef(SpreadsheetCell cell, string root)
        {
            bool Error = false;
            string CellName = CellToString(cell);

            if (cell.VarList != null)
            {
                foreach (string ReferencedCell in cell.VarList) // Will check all variables in the cell's variable list
                {
                    if (null == StringToCell(ReferencedCell)) // If the parsing of the referenced cell failed, the error should be bad_ref not circ_ref
                        return false;
                    //else if (cell.Value.Contains("CIRC_REF"))
                    //    root = CellToString(cell);
                    else if (CellName != root && ReferencedCell == root) // If the loop came back around to the root value while checking, then a circular reference has been found
                        return true;
                    else
                        Error = CheckCircularRef(StringToCell(ReferencedCell), root); // Continues checks for circular referencing
                }
            }

            return Error;
        }

        private HashSet<string> GetVarNames(string exp)
        {
            string[] all = exp.Split(new char[8] { '*', '/', '+', '-', '=', ' ', '(', ')'}); // Splits the expression into contants and variables

            HashSet<string> variables = new HashSet<string>();

            foreach(string part in all)
            {
                if (part != "") // Ignores blank space
                {
                    if (Char.IsLetter(part[0])) // Ensures that the first character is an uppercase char
                        variables.Add(part);
                }
            }

            return variables;
        }

        internal SpreadsheetCell StringToCell(string name)
        {
            char temp = name[0];
            int j = temp - 65; // converts the letter into the index (Ex: Column A -> 0)

            if (j < 0 || j > this.width)
                return null;

            name = name.Remove(0, 1);
            int i = 0;
            if (Int32.TryParse(name, out i)) // Finds Row
            {
                if (i < 0 || i > this.height)
                    return null;
            }
            else
                return null;

            SpreadsheetCell from = cellArr[i - 1, j]; // Finds the cell given the parsed row and column number

            return from;
        }

        internal string CellToString(SpreadsheetCell from)
        {
            char column = (char)(from.ColumnIndex + 65); // Converts column index to an uppercase char
            int row = from.RowIndex + 1; // corrects the row offset

            return column.ToString() + row.ToString(); // Returns the string version of the cell
        }

        public void PushUndo(CmdCollection PropertyCollection)
        {
            UndoStack.Push(PropertyCollection); // Pushes the restoring collection on the undo stack
            RedoStack.Clear(); // Clears the redo stack -> ensures that operations cannot be restored after preceding undo operations have been restored
        }

        public void PopUndo()
        {
            if (UndoStack.Count > 0) // ensures that there is something to be popped
            {
                CmdCollection from = UndoStack.Pop(); // pops the top list off of the undo stack -> all of the contents will be undone
                /* Executes all commands in the popped list
                 * Pushes the inverse of the popped actions onto the redo stack */
                RedoStack.Push(new CmdCollection(from.ExecuteAll()));
            }
        }

        public void PopRedo()
        {
            if (RedoStack.Count > 0) // ensures that there is something to be popped
            {
                CmdCollection from = RedoStack.Pop();

                /* Executes all commands in the popped list
                 * Pushes the inverse of the popped actions onto the redo stack */
                UndoStack.Push(new CmdCollection(from.ExecuteAll()));
            }
        }

        public string PeekUndo
        {
            /* Will return the type of the command that is on the top of this stack -> used for display purposes */
            get { return UndoStack.Peek().GetType; }
        }

        public string PeekRedo
        {
            /* Will return the type of the command that is on the top of this stack -> used for display purposes */
            get { return RedoStack.Peek().GetType; }
        }

        public int UndoCount
        {
            /* Returns the count of the stack -> ensures that the stack can be popped from in the UI */
            get { return UndoStack.Count; }
        }

        public int RedoCount
        {
            /* Returns the count of the stack -> ensures that the stack can be popped from in the UI */
            get { return RedoStack.Count; }
        }

        internal void ClearUndoRedoSystem()
        {
            this.UndoStack.Clear();
            this.RedoStack.Clear();
        }
    }

    public class CmdCollection
    {
        private List<IUndoRedoCmd> cmds; // Will store various restoring commands

        public CmdCollection()
        {
            /* constructor used to initialize a new list of IUndoRedoCmds */
            cmds = new List<IUndoRedoCmd>();
        }

        public CmdCollection(List<IUndoRedoCmd> from)
        {
            /* Will copy over an entire list into this object's list */
            cmds = from;
        }

        public CmdCollection(CmdCollection from)
        {
            /* Copy constructor
             * simply copies over the list in 'from' */
            this.cmds = from.cmds;
        }

        public CmdCollection(IUndoRedoCmd Changed)
        {
            /* Constructor for passing in a single operation to restore to */
            this.cmds = new List<IUndoRedoCmd>();
            this.cmds.Add(Changed);
        }

        public CmdCollection ExecuteAll()
        {
            CmdCollection inverses = new CmdCollection(); // Will contain the opposite operations to the ones that are currently contained in this object

            foreach (IUndoRedoCmd cmd in this.cmds)
            {
                /* Adds the inverse operations to a NEW CmdCollection
                 * This will be pushed onto the opposite stack from where the current operation came from */
                inverses.cmds.Add(cmd.Execute()); 
            }

            return inverses;
        }

        public void Add(IUndoRedoCmd toAdd)
        {
            this.cmds.Add(toAdd);
        }

        public new string GetType
        {
            get { return cmds[0].PropertyName(); }
        }
    }

    public class SaveLoad
    {
        public event PropertyChangedEventHandler SavedPropertyChanged = delegate { };
        private bool isSaved;
        private string currentFile;

        public SaveLoad()
        {
            this.currentFile = ""; // Every "new" form comes in without a file name attached to it
            this.isSaved = true; // a new form is has no changes to it, so it is already saved
        }

        public SaveLoad(string file)
        {
            currentFile = file;
            isSaved = true;
        }

        public string FileName
        {
            get { return currentFile; }
            set
            {
                if (currentFile == value) { return; }

                currentFile = value;
                SavedPropertyChanged(this, new PropertyChangedEventArgs("FileName")); // Will notify the subscribed objects that the file has changed
            }
        }

        public bool Saved
        {
            get { return isSaved; }

            set
            {
                if (isSaved == value) { return; }

                isSaved = value;
                SavedPropertyChanged(this, new PropertyChangedEventArgs("Saved")); // Will notify the subscribed objects that the saved bool has changed
            }
        }

        public void Save(string ToFile, Spreadsheet Sender)
        {
            using (StreamWriter WriteStream = new StreamWriter(ToFile)) // Uses a StreamWriter -> can be changed to any stream
            {
                try
                {
                    Save_Pvt(WriteStream, Sender); // Saves to the stream using XML Writer
                    this.Saved = true; // If saving was successful, the save value is set to true
                    this.FileName = ToFile; // current file name is updated
                }
                catch(Exception e)
                {
                    throw e;
                }
            }
        }

        private bool Save_Pvt(TextWriter ToStream, Spreadsheet Sender)
        {
            XmlWriter xmlWrite = null;

            try
            {
                xmlWrite = XmlWriter.Create(ToStream); // initializes the XML Writer with the passed stream
            }
            catch (ArgumentException e)
            {
                throw e;
            }

            if (xmlWrite != null) // Ensures the stream is usable
            {
                using (xmlWrite)
                {
                    xmlWrite.WriteStartDocument(); // Places start tag
                    xmlWrite.WriteStartElement("Spreadsheet"); // Places the root node -> the spreadsheet itself

                    for (int i = 0; i < Sender.RowCount; i++) // iterates through every cell in the spreadsheet
                    {
                        for (int j = 0; j < Sender.ColumnCount; j++)
                        {
                            SpreadsheetCell from = Sender[i, j]; // grabs a cell

                            if (!Sender[i, j].IsDefault()) // Only saves those cells that have been altered
                            {
                                xmlWrite.WriteStartElement("Cell"); // Creates a cell start tag
                                xmlWrite.WriteAttributeString("Name", Sender.CellToString(from)); // Gives the cell a NAME attribute

                                if (from.Text != "") // No point in saving an empty string
                                {
                                    xmlWrite.WriteElementString("Text", from.Text); // sets the TEXT from the cell
                                    /* Cell's VALUE does NOT have to be saved
                                     * Every cell's text will be re-computed when it is loaded in
                                     * order does NOT matter, cell's that have references to it will simply be updated when it's their turn
                                     */
                                }
                                xmlWrite.WriteElementString("BGColor", from.BGColor.ToString()); // writes the BGColor element inside of the cell

                                xmlWrite.WriteEndElement(); // Ends the cell block
                            }
                        }
                    }

                    xmlWrite.WriteEndElement(); // Ends the Spreadsheet block
                    xmlWrite.WriteEndDocument(); // Ends file
                }

                return true;
            }

            return false;
        }

        public void Load(string FromFile, Spreadsheet Sender)
        {
            using (StreamReader ReadStream = new StreamReader(FromFile)) // Uses a StreamReader -> can be changed to any stream
            {
                try
                {
                    Load_Pvt(ReadStream, Sender); // attempts to load from the specified stream
                    Sender.ClearUndoRedoSystem(); // Clears the undo/redo system
                    this.FileName = FromFile; // updates the file name to the file that was just loaded in (allows to save to it again without prompting)
                }
                catch (XmlException e)
                {
                    throw e;
                }
            }
        }

        private bool Load_Pvt(TextReader FromStream, Spreadsheet Sender)
        {
            XmlReader xmlRead = XmlReader.Create(FromStream); // initializes the XMLReader with the passed stream

            if (xmlRead != null)
            {
                using (xmlRead) // opens the stream
                {
                    SpreadsheetCell to = null; // changes to the cell that is pulled from the .xml file
                    while (!xmlRead.EOF) // while there is still information to be read
                    {
                        if (xmlRead.IsStartElement()) // ensures that the reader is at a start element
                        {
                            switch (xmlRead.Name) // checks various names -> skips over unknown names
                            {
                                case "Cell":
                                    xmlRead.MoveToFirstAttribute(); // moves to the name attribute
                                    string cell = xmlRead.Value;
                                    to = Sender.StringToCell(cell); // will grab the cell from the spreadsheet using this name/attribute
                                    xmlRead.Read(); 
                                    break;
                                case "Text":
                                    Sender[to.RowIndex, to.ColumnIndex].Text = xmlRead.ReadElementContentAsString(); // sets Text
                                    break;
                                case "BGColor":
                                    Sender[to.RowIndex, to.ColumnIndex].BGColor = xmlRead.ReadElementContentAsInt(); // sets BGColor
                                    break;
                                default: // simply reads over the unknown element -> also skips spreadsheet name (nothing is done here at the moment)
                                    xmlRead.Read();
                                    break;
                            }
                        }
                        else
                            xmlRead.Read(); // simply passed anything that is not a start element
                    }
                }

                return true;
            }

            return false;
        }
    }
}

namespace UndoRedoSystem
{
    using SpreadsheetEngine;
    public interface IUndoRedoCmd
    {
        IUndoRedoCmd Execute(); // Every Property will inherit this functionality -> allows the spreadsheet to return to this point

        string PropertyName(); // Every inheriting class must have a function that returns the name of the property to restore
    }

    public class RestoreText : IUndoRedoCmd
    {
        private SpreadsheetCell Current;
        private string CellText;

        public RestoreText(SpreadsheetCell from, string text)
        {
            this.Current = from;
            this.CellText = text;
        }

        public IUndoRedoCmd Execute()
        {
            var inverse = new RestoreText(Current, Current.Text); // Creates the inverse class -> restores to this
            Current.Text = CellText; // restores the cell's text to the contained text
            return inverse;
        }

        /* PropertyName:
         * Returns the name of the property in string format*/
        public string PropertyName()
        {
            return "Text";
        }
    }

    public class RestoreBGColor : IUndoRedoCmd
    {
        private SpreadsheetCell Current;
        private int ARBGColor;

        public RestoreBGColor(SpreadsheetCell from, int color)
        {
            this.Current = from;
            this.ARBGColor = color;
        }

        public IUndoRedoCmd Execute()
        {
            var inverse = new RestoreBGColor(Current, Current.BGColor); // Creates the inverse class -> restores to this
            Current.BGColor = ARBGColor; // restores the cell's color to the contained color
            return inverse;
        }

        /* PropertyName:
        * Returns the name of the property in string format*/
        public string PropertyName()
        {
            return "Cell Color";
        }
    }
}

namespace ExpressionTree
{
    public class ExpTree
    {
        private static Dictionary<string, double> mainDict;
        private Node mRoot;
        private string mExp;


        /*All Node Classes*/
        private abstract class Node
        {
            abstract public double Execute(); 
            /* Every node must have an execute function That returns a double
             * OpNode -> returns result (as a double) of its two subtrees
             * ConstNode -> Simply returns its value
             * VarNode -> Looks up its value in the dictionary and returns that value */
        }

        private class ConstNode : Node
        {
            /*This node will simply represent a constant -> (ex: 1, 2, 120.57, .5)*/
            private readonly double mVal;

            public ConstNode(double newVal) // Constructor -> simply places the new value into the mVal memVar
            {
                mVal = newVal;
            }

            public double Value { get { return mVal; } }

            public override double Execute() // When called upon, will simply return its value
            {
                return this.Value;
            }
        }

        private class VarNode : Node
        {
            /* This node will simply represent a variable -> (ex: X, Y, A1, F7)
             * When executed, the node will look up its value in the tree's dictionary */

            private readonly string mName;

            public VarNode(string newName) // Contains the string that will be used to lookup the variable's value in the dictionary
            {
                mName = newName;
            }

            public string Value { get { return mName; } }

            public override double Execute()
            {
                double dictVal;
                if (mainDict.TryGetValue(this.Value, out dictVal)) // looks up the variables value in the dictionary
                    return dictVal; // returns the found value
                else
                    return 0.0; // will return 0.0 if no value is found
            }
        }

        private class OpNode : Node
        {
            /* Contains:
             * String representing the operator that will be applied to its left and right children
             * a left sub tree of type Node
             * a right sub tree of type Node 
             * returns the operation of its left and right subtrees when executed*/

            private readonly string mOperator;
            public Node left;
            public Node right;

            public OpNode(string newOp, Node childLeft, Node childRight)  
            {
                mOperator = newOp;
                left = childLeft; // Only operator nodes have children
                right = childRight;
            }

            public string Value { get { return mOperator; } }

            public override double Execute()
            {
                double result = 0.0;

                if (this.left != null)
                {
                    result = GetTreeValue(this.left); // Gathers the results of the left subtree
                }
                if (this.right != null)
                {
                    double temp = GetTreeValue(this.right); // Gathers the results of the right subtree
                    try
                    {
                        PerformOperation(this, temp, ref result); // Combines the left and the right subtrees using the correct operation
                    }
                    catch (DivideByZeroException e) // Dividing by zero occured
                    {
                        throw e;
                    }
                }
                return result;
            }
        }

        public ExpTree()
        {
            mainDict = new Dictionary<string, double>();
            mExp = "((A1*B1)-(A1+B1))"; // Default Expression
            mRoot = Build(mExp); // When instansiated, will automatically build the expression tree
        }

        public ExpTree(string expression)
        {
            mExp = expression;
            mRoot = Build(mExp); // Automatically builds the expression tree when a new expression is entered
            mainDict = new Dictionary<string, double>(); // Creates a new dictionary 
        }

        public string Expression
        {
            get { return mExp; } // Returns the expression

            set
            {
                mExp = value;
                mRoot = Build(mExp); // Builds the expression tree when a new expression is entered
                mainDict = new Dictionary<string, double>(); // Creates a new dictionary (erases old variable values)
            }
        }

        public void SetVar(string varName, double varValue)
        {
            if (mainDict != null) // Ensures that the dictionary has been initialized
                mainDict[varName] = varValue; // Places the value into the dictionary
        }

        public double Eval()
        {
            double result = 0.0;

            if (mRoot == null) // Building the tree returned null -> *Invalid expression*
                throw new NullReferenceException();

            try
            {
                result = GetTreeValue(mRoot); // Will evaluate the result using the pre-built epxression tree
            }
            catch (DivideByZeroException e)
            {
                throw e; // NaN
            }

            return result;
        }

        private Node Build(string exp)
        {
            int check = inParenthesis(exp); // Checks if the entire expression/subexpression is enclosed in parenthesis
            if (check != -1) // Valid Expression
            {
                if (check == 1) // Expression was valid, but inside parenthesis
                    return Build(exp.Substring(1, exp.Length - 2)); // Recursively takes off parenthesis
                else // Valid and NOT in parenthesis
                {
                    int opIdx = GetNextOp(exp); // gets the index of the next operator
                    Node left = null, right = null;

                    if (opIdx < 0) // Has no operators / none found (out of +, -, /, *)
                    {
                        double expVal;
                        if (Double.TryParse(exp, out expVal))
                            return new ConstNode(expVal);
                        else // Find the variable's value
                            return new VarNode(exp);
                    }
                    else
                    {
                        left = Build(exp.Substring(0, opIdx)); // passes in the substring from the beginning of the string to the index before the next operator
                        right = Build(exp.Substring(opIdx + 1)); // passes in the substring from the index of the next operator, plus one, on...
                    }

                    return new OpNode(exp[opIdx].ToString(), left, right); // returns the tree/subtree with all of the children
                }
            }
            else
                return null;
        }

        private int inParenthesis(string s)
        {
            int counter = 0;
            bool inParen = false;

            if (s == "") // Empty Expression
                return -1;
            string sub;
            if (s[0] == '(') // if the first character is a '(', there is the possibility for the entire exp to be in parenthesis
            {
                inParen = true;
                sub = s.Substring(1);
                counter++;
            }
            else
                sub = s;

            foreach (char c in sub)
            {
                if (counter < 0) // automatically an invalid expression
                    return -1;
                if (inParen && counter == 0) // a balanced number of '(' and ')' has been reached before the end of the exp -> exp not surrounded by paranthesis
                    inParen = false; // Will continue to check for an invalid expression even though this has become false
                if (c == '(')
                    counter++;
                if (c == ')')
                    counter--;
            }

            if (counter > 0) // Invalid expression
                return -1;
            else if (inParen) // In Parenthesis
                return 1;
            else // Not in parenthesis
                return 0;
        }

        private int GetNextOp(string exp)
        {
            int counter = 0;
            int lowestOp = -1;
            for (int i = exp.Length - 1; i >= 0; i--) // reads the string from right to left
            {
                if (exp[i] == '(')
                    counter++;
                else if (exp[i] == ')')
                    counter--;
                else if (counter == 0 && IsOperator(exp[i]))
                {
                    if (exp[i] == '+' || exp[i] == '-')
                        return i; // immediately returns if either a '+' or a '-' is found -> cannot have a lower precedence than that
                    else if (-1 == lowestOp) // Will only update at the first occurance of a '*' or a '/' in this case
                        lowestOp = i;
                }
            }

            return lowestOp; // signifies that no operator was found in the passed string
        }

        static private bool IsOperator(char c)
        {
            if (c == '*')
                return true;
            if (c == '/')
                return true;
            if (c == '+')
                return true;
            if (c == '-')
                return true;

            return false; // no operator found (out of this list)
        }

        static private void PerformOperation(OpNode operation, double val, ref double result) // Performs the operation contained in the opnode
        {
            if (operation.Value == "+") // Addition
                result += val;
            else if (operation.Value == "-") // Subtraction
                result -= val;
            else if (operation.Value == "*") // Multiplication
                result *= val;
            else // Division
            {
                if (val != 0) // ensures no dividing by zero
                    result /= val;
                else // otherwise result will not change
                    throw new DivideByZeroException();
            }
        }

        static private double GetTreeValue(Node root)
        {
            double result = 0.0;

            if (root != null) // Ensures that there are nodes in the tree
                result = root.Execute();

            return result;
        }
    } 
}