using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Square_ExtractData_CreateTable
{

    // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public partial class Layers
    {

        private LayersLayer[] layerField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Layer")]
        public LayersLayer[] Layer
        {
            get
            {
                return this.layerField;
            }
            set
            {
                this.layerField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class LayersLayer
    {

        private string sNoField;

        private string descriptionField;

        private string nameField;

        private byte redField;

        private byte greenField;

        private byte blueField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string SNo
        {
            get
            {
                return this.sNoField;
            }
            set
            {
                this.sNoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Description
        {
            get
            {
                return this.descriptionField;
            }
            set
            {
                this.descriptionField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte Red
        {
            get
            {
                return this.redField;
            }
            set
            {
                this.redField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte Green
        {
            get
            {
                return this.greenField;
            }
            set
            {
                this.greenField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte Blue
        {
            get
            {
                return this.blueField;
            }
            set
            {
                this.blueField = value;
            }
        }
    }
}