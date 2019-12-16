using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;

namespace MisClasesFuncionesC
{
    /// <summary>
	/// Descripción breve de MiTextBox.
	/// </summary>
	[ToolboxBitmap(typeof(TextBox))]
    public class MiTextBox : TextBox
    {

        #region VARIABLES de la Instancia

        public System.Drawing.Color BackColor_Nuevo;
        private bool swAplicarFormato = false;
        private bool swEvitarOnTextChangedRecursivo = false;

        private bool _UsarEnterComoTab = true;
        private bool _NumericoSolo = false;
        private bool _NumeroSolo = false;
        private string _Formato = string.Empty;
        private string _ValorReal = string.Empty;
        private bool _SeleccionarTodoCuandoClick = false;
        private bool _TextoVacioNumPermitido = false;
        private const string MASCARA_NUMERICO = "1234567890,";
        private const string MASCARA_NUMERO = "1234567890";

        #endregion

        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public MiTextBox(System.ComponentModel.IContainer container)
        {
            ///
            /// Requerido para la compatibilidad con el Diseñador de composiciones de clases Windows.Forms
            ///
            container.Add(this);
            InitializeComponent();
        }

        #region PROPIEDADES

        [Category("miPropiedades")]
        [Description("Indica si el TextBox es campo numerio (solo numeros y una coma/punto)")]
        [MergableProperty(false)]
        public bool NumericoSolo
        {
            get { return this._NumericoSolo; }
            set { this._NumericoSolo = value; }
        }

        [Category("miPropiedades")]
        [Description("Indica que el TextBox solo admite Numeros")]
        [MergableProperty(false)]
        public bool NumeroSolo
        {
            get
            { return this._NumeroSolo; }
            set { this._NumeroSolo = value; }
        }

        [Category("miPropiedades")]
        [Description("Formato a aplicar al TextBox. (Solo Campos Numericos y Numeros)")]
        public string Formato
        {
            get { return _Formato; }
            set
            {
                _Formato = value.Trim();
                if (this._Formato == string.Empty)
                { this.swAplicarFormato = false; }
                else
                { this.swAplicarFormato = true; }
            }
        }

        /// <summary>
        /// Valor real del TextBox sin Formato Aplicado
        /// </summary>
        public string ValorReal
        {
            get { return this._ValorReal; }
        }

        [Category("miPropiedades")]
        [Description("Controla si la Tecla ENTER actua como TAB.")]
        [MergableProperty(true)]
        public bool UsarEnterComoTab
        {
            get { return this._UsarEnterComoTab; }
            set { this._UsarEnterComoTab = value; }
        }

        [Category("miPropiedades")]
        [Description("Controla si se selecciona todo el Texto cunado se entra via Click.")]
        [MergableProperty(false)]
        public bool SeleccionarTodoCuandoClick
        {
            get { return this._SeleccionarTodoCuandoClick; }
            set { this._SeleccionarTodoCuandoClick = value; }
        }

        [Category("miPropiedades")]
        [Description("Indica si en Campos Numericos (NumericoSolo y NumeroSolo) se permite Texto Vacio (String.Empty).")]
        [MergableProperty(false)]
        public bool TextoVacioNumPermitido
        {
            get { return this._TextoVacioNumPermitido; }
            set { this._TextoVacioNumPermitido = value; }
        }

        #endregion

        public MiTextBox()
        {
            ///
            /// Requerido para la compatibilidad con el Diseñador de composiciones de clases Windows.Forms
            ///
            InitializeComponent();

            //
            // TODO: agregar código de constructor después de llamar a InitializeComponent
            //
            this._ValorReal = this.Text;
            this._Formato = string.Empty;
        }

        #region Manejadores de EVENTOS

        protected override void OnEnter(EventArgs e)
        {
            if (!this.ReadOnly)
            {
                this.BackColor_Nuevo = this.BackColor;
                this.BackColor = Color.Gold;
            }

            if (swAplicarFormato & (this.NumericoSolo | this.NumeroSolo))
            {
                this.swEvitarOnTextChangedRecursivo = true;
                this.Text = this._ValorReal;
                this.swEvitarOnTextChangedRecursivo = false;
            }

            this.SelectAll();

            base.OnEnter(e);
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);

            if (this._SeleccionarTodoCuandoClick)
            {
                this.SelectAll();
            }

        }

        protected override void OnTextChanged(EventArgs e)
        {
            if (this.swEvitarOnTextChangedRecursivo)
            {
                return;
            }

            if (this._NumericoSolo | this._NumeroSolo)
            {
                if (this.Text == string.Empty & !this._TextoVacioNumPermitido)
                {
                    this.swEvitarOnTextChangedRecursivo = true;
                    this.Text = "0";
                    this.swEvitarOnTextChangedRecursivo = false;
                    SendKeys.Send("{RIGHT}");
                }
                else if (this.Text == ",")
                {
                    this.swEvitarOnTextChangedRecursivo = true;
                    this.Text = "0,";
                    this.swEvitarOnTextChangedRecursivo = false;
                    SendKeys.Send("{RIGHT}");
                    SendKeys.Send("{RIGHT}");
                }
            }

            this._ValorReal = this.Text;
            if (!this.ContainsFocus & swAplicarFormato & (this.NumericoSolo | this.NumeroSolo))
            {
                this.swEvitarOnTextChangedRecursivo = true;
                if (this.Text != string.Empty)
                {
                    this.Text = Convert.ToDecimal(this._ValorReal).ToString(this._Formato);
                }
                this.swEvitarOnTextChangedRecursivo = false;
            }

            base.OnTextChanged(e);
        }

        protected override void OnLeave(EventArgs e)
        {
            if (!this.ReadOnly)
            {
                this.BackColor = this.BackColor_Nuevo;
                // this.BackColor = Color.Empty;
            }

            if (this.swAplicarFormato & (this.NumericoSolo | this.NumeroSolo))
            {
                if (this.Text != string.Empty)
                {
                    this.swEvitarOnTextChangedRecursivo = true;
                    this.Text = Convert.ToDecimal(this._ValorReal).ToString(this._Formato);
                    this.swEvitarOnTextChangedRecursivo = false;
                }
            }

            base.OnLeave(e);
        }

        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            base.OnKeyPress(e);

            if (this._NumericoSolo)
            {
                switch (e.KeyChar)
                {
                    case (char)8:
                        break;
                    case (char)13:
                        break;
                    default:
                        if (e.KeyChar == ',')
                        {
                            if (this.TextLength != this.SelectionLength & this.Text.IndexOf(',') >= 0)
                            {
                                MiComun.MessageBeep(MiComun.MessageBeepType.Default);
                                e.Handled = true;
                            }
                        }
                        else if (MASCARA_NUMERICO.IndexOf(e.KeyChar) < 0)
                        {
                            MiComun.MessageBeep(MiComun.MessageBeepType.Default);
                            e.Handled = true;
                        }
                        break;
                }
            }
            if (this._NumeroSolo)
            {
                switch (e.KeyChar)
                {
                    case (char)8:
                        break;
                    case (char)13:
                        break;
                    default:
                        if (MASCARA_NUMERO.IndexOf(e.KeyChar) < 0)
                        {
                            MiComun.MessageBeep(MiComun.MessageBeepType.Default);
                            e.Handled = true;
                        }
                        break;
                }
            }
        }

        protected override void OnKeyDown(System.Windows.Forms.KeyEventArgs e)
        {
            if (this._NumericoSolo)
            {
                if (e.KeyCode == Keys.Decimal)
                {
                    SendKeys.Send(",");
                    e.Handled = true;
                }
            }

            base.OnKeyDown(e);
        }

        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, System.Windows.Forms.Keys keyData)
        {
            if (_UsarEnterComoTab & keyData == Keys.Enter)
            {
                SendKeys.Send("{Tab}");
                return true;
            }
            else
            {
                return base.ProcessCmdKey(ref msg, keyData);
            }

        }

        public void Vaciar()
        {
            this.swEvitarOnTextChangedRecursivo = true;
            this.Text = string.Empty;
            this.swEvitarOnTextChangedRecursivo = false;
        }

        /// <summary> 
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #endregion

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {

        }

        #endregion

    }
}
