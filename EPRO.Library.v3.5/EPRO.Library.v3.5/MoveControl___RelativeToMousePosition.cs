using System;
using System.Drawing;
using System.Windows.Forms;

namespace ELibrary.Standard
{

    /// <summary>
/// Helps to move anther control relative to mouse position after clicking the invoke. Works on left click and dragging
/// </summary>
/// <remarks></remarks>
    public class MoveControl___RelativeToMousePosition : IDisposable
    {







        #region Properties


        private Control ______Control_to_InvokeDrag;
        private Control ______Control_to_Relocate;

        public bool IsDisposed
        {
            get
            {
                return disposedValue;
            }
        }

        private Point _____MouseLocation____OnBeginMove;
        private Point _____ControlToMove__Location__OnBeginMove;
        private bool _____KeepMovingControl;



        #endregion




        #region Constructors


        /// <summary>
    /// Helps to drag control
    /// </summary>
    /// <param name="p______Control_to_InvokeDrag">Control that will be mouse clicked to move</param>
    /// <param name="p______Control_to_Relocate">Control that will be moved</param>
    /// <remarks></remarks>
        public MoveControl___RelativeToMousePosition(Control p______Control_to_InvokeDrag, Control p______Control_to_Relocate)
        {
            if (p______Control_to_InvokeDrag is null || p______Control_to_Relocate is null || p______Control_to_InvokeDrag.IsDisposed || p______Control_to_Relocate.IsDisposed)
                throw new InvalidOperationException("You are calling this class with wrong or disposed parameters");
            ______Control_to_InvokeDrag = p______Control_to_InvokeDrag;
            ______Control_to_Relocate = p______Control_to_Relocate;
            ______Control_to_InvokeDrag.MouseDown += ______Control_to_InvokeDrag_MouseDown;
            ______Control_to_InvokeDrag.MouseUp += ______Control_to_InvokeDrag_MouseUp;
            ______Control_to_InvokeDrag.MouseMove += ______Control_to_InvokeDrag_MouseMove;
        }




        #endregion





        #region Events

        public event MovedEventHandler Moved;

        public delegate void MovedEventHandler(MoveControl___RelativeToMousePosition pSender);

        #endregion



















        #region IDisposable Support
        private bool disposedValue; // To detect redundant calls

        // IDisposable
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).

                    if (______Control_to_InvokeDrag is object && !______Control_to_InvokeDrag.IsDisposed)
                    {
                        ______Control_to_InvokeDrag.MouseDown -= ______Control_to_InvokeDrag_MouseDown;
                        ______Control_to_InvokeDrag.MouseUp -= ______Control_to_InvokeDrag_MouseUp;
                        ______Control_to_InvokeDrag.MouseMove -= ______Control_to_InvokeDrag_MouseMove;
                    }
                }

                // TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                // TODO: set large fields to null.
            }

            disposedValue = true;
        }

        // TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        // Protected Overrides Sub Finalize()
        // ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        // Dispose(False)
        // MyBase.Finalize()
        // End Sub

        // This code added by Visual Basic to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion





        #region Methods




        private void ______Control_to_InvokeDrag_MouseDown(object sender, MouseEventArgs e)
        {
            _____MouseLocation____OnBeginMove = Control.MousePosition;
            _____ControlToMove__Location__OnBeginMove = ______Control_to_Relocate.Location;
            _____KeepMovingControl = true;
        }

        private void ______Control_to_InvokeDrag_MouseUp(object sender, MouseEventArgs e)
        {
            _____KeepMovingControl = false;
        }

        private void ______Control_to_InvokeDrag_MouseMove(object sender, MouseEventArgs e)
        {
            var pNewMousePoint = Control.MousePosition;
            if (e.Button == MouseButtons.Left && _____KeepMovingControl)
            {
                ______Control_to_Relocate.Location = new Point(_____ControlToMove__Location__OnBeginMove.X + (pNewMousePoint.X - _____MouseLocation____OnBeginMove.X), _____ControlToMove__Location__OnBeginMove.Y + (pNewMousePoint.Y - _____MouseLocation____OnBeginMove.Y));
                Moved?.Invoke(this);
            }
        }




        #endregion



    }
}