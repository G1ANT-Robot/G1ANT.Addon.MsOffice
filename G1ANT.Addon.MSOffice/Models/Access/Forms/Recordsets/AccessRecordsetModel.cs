using System;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Fields;
using Microsoft.Office.Interop.Access.Dao;

namespace G1ANT.Addon.MSOffice.Models.Access.Forms.Recordsets
{
    internal class AccessRecordsetModel : INameModel
    {
        private Recordset recordset;

        public Lazy<AccessDaoFieldCollectionModel> Fields { get; }
        public EditModeEnum EditMode => (EditModeEnum)recordset.EditMode;
        public string Name => recordset.Name;
        public RecordsetTypeEnum Type => (RecordsetTypeEnum)recordset.Type;

        public DateTime? LastUpdated { get { try { return recordset.LastUpdated; } catch { return null; } } }
        public DateTime? DateCreated { get { try { return recordset.DateCreated; } catch { return null; } } }

        public int AbsolutePosition { get => recordset.AbsolutePosition; set => recordset.AbsolutePosition = value; }

        public int RecordCount => recordset.RecordCount;

        public RecordStatusEnum RecordStatus { get { try { return (RecordStatusEnum)recordset.RecordStatus; } catch { return (RecordStatusEnum)(-1); } } }

        public string Sort { get => recordset.Sort; set => recordset.Sort = value; }

        public string Filter { get => recordset.Filter; set => recordset.Filter = value; }

        public AccessConnectionModel Connection
        {
            get { try { return new AccessConnectionModel(recordset.Connection); } catch { return null; } }
            set => recordset.Connection = value.Connection;
        }

        public Array LastModified => recordset.LastModified;

        //public dynamic Index { get; }

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



        public AccessRecordsetModel(Recordset recordset)
        {
            this.recordset = recordset;

            Fields = new Lazy<AccessDaoFieldCollectionModel>(() => new AccessDaoFieldCollectionModel(recordset.Fields));
            
            //Index = recordset.Index;
        }


        public override string ToString() => Name;
    }
}
