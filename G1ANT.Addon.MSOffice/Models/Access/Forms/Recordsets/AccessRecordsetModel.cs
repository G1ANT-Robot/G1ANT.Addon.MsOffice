using System;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Fields;
using Microsoft.Office.Interop.Access.Dao;

namespace G1ANT.Addon.MSOffice.Models.Access.Forms.Recordsets
{
    internal class AccessRecordsetModel : INameModel
    {
        private Recordset recordset;

        public AccessDaoFieldCollectionModel Fields { get; }
        public string EditMode { get; }
        public string Name { get; }
        public Array LastModified { get; }

        //public dynamic Index { get; }
        public int AbsolutePosition { get => recordset.AbsolutePosition; set => recordset.AbsolutePosition = value; }

        public int RecordCount { get => recordset.RecordCount; }

        public string Sort { get => recordset.Sort; set => recordset.Sort = value; }

        public string Filter { get => recordset.Filter; set => recordset.Filter = value; }

        public AccessConnectionModel Connection
        {
            get { try { return new AccessConnectionModel(recordset.Connection); } catch { return null; } }
            set => recordset.Connection = value.Connection;
        }

        /// <summary>
        /// Sets or returns a value indicating the type of locking that is in effect while editing.
        /// True: Default. Pessimistic locking is in effect.The page containing the record you're editing is locked as soon as you call the Edit method.
        /// False: Optimistic locking is in effect for editing.The page containing the record is not locked until the Update method is executed.
        /// </summary>
        public bool LockEdits { get => recordset.LockEdits; set => recordset.LockEdits = value; }

        public bool NoMatch { get => recordset.NoMatch; }

        /// <summary>Beginning of a File</summary>
        public bool BOF => recordset.BOF;
        
        /// <summary>End of a File</summary>
        public bool EOF => recordset.EOF;

        public dynamic DateCreated { get; }

        public AccessRecordsetModel(Recordset recordset)
        {
            this.recordset = recordset;

            Fields = new AccessDaoFieldCollectionModel(recordset.Fields);

            EditMode = ((EditModeEnum)recordset.EditMode).ToString();

            Name = recordset.Name;

            //Index = recordset.Index;

            LastModified = recordset.LastModified;
            try { DateCreated = recordset.DateCreated; } catch { }
        }


        public override string ToString() => Name;
    }
}
