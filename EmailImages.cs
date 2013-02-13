using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DM.Core.Net
{
    public class EmailImages : IList<EmailImage>, IDisposable
    {
        List<EmailImage> __List;

        public EmailImages()
            : base()
        {
            __List = new List<EmailImage>();
        }

        #region IList<EmailImage> Members

        public int IndexOf(EmailImage item) { return __List.IndexOf(item); }
        public void Insert(int index, EmailImage item) { __List.Insert(index, item); }
        public void RemoveAt(int index) { __List.RemoveAt(index); }

        public EmailImage this[int index]
        {
            get { return __List[index]; }
            set { __List[index] = value; }
        }

        #endregion

        #region ICollection<EmailImage> Members

        public void Add(EmailImage item) { __List.Add(item); }
        public void Clear() { __List.Clear(); }
        public bool Contains(EmailImage item) { return __List.Contains(item); }
        public void CopyTo(EmailImage[] array, int arrayIndex) { __List.CopyTo(array, arrayIndex); }
        public bool Remove(EmailImage item) { return __List.Remove(item); }

        public int Count
        {
            get { return __List.Count; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }



        #endregion

        #region IEnumerable<EmailImage> Members

        public IEnumerator<EmailImage> GetEnumerator() { return __List.GetEnumerator(); }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }

        #endregion

        public void Dispose()
        {

            try
            {
                foreach (EmailImage img in this.__List)
                    img.Dispose();
            }
            catch { /*none*/ }
            __List.Clear();
            __List = null;

        }
    }
}
