using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace kuba.extensions.vs.aliaser
{
    internal class AliaserService : IAliaserService, IDisposable
    {
        /// <summary>
        /// Dictionary with Type instances and alias
        /// Key = Instance
        /// Value = alias
        /// </summary>
        private readonly Dictionary<String, String> _typeDictionary;

        public AliaserService()
        {
            this._typeDictionary = this.GetResourceStringDictionary();
        }

        public String GetTransformedText(String selectedText, TransformationFlow flow)
        {
            Boolean wasChangesMade = false;
            return this.GetTransformedText(selectedText, flow, out wasChangesMade);
        }

        public String GetTransformedText(String selectedText, TransformationFlow flow, out Boolean wasChangesMade)
        {
            wasChangesMade = false;

            if (String.IsNullOrEmpty(selectedText) || this._typeDictionary == null || this._typeDictionary.Count == 0)
                return selectedText;
            else
            {
                String resultText = selectedText;
                foreach (KeyValuePair<String, String> item in this._typeDictionary)
                {
                    KeyValuePair<String, String> transPair = item;

                    if (flow == TransformationFlow.TO_INSTANCE)
                        transPair = new KeyValuePair<String, String>(transPair.Value, transPair.Key);
                    
                    if (resultText.Contains(transPair.Key))
                    {
                        resultText = resultText.Replace(transPair.Key, transPair.Value);
                        wasChangesMade = true;
                    }
                }

                return resultText;
            }
        }

        private Dictionary<String, String> GetResourceStringDictionary()
        {
            Dictionary<String, String> result = new Dictionary<String, String>();

            ResourceManager rm = Properties.Resources.ResourceManager;

            ResourceSet rs = rm.GetResourceSet(new CultureInfo("en-US"), true, true);

            if (rs != null)
            {
                IDictionaryEnumerator de = rs.GetEnumerator();
                while (de.MoveNext())
                {
                    if (de.Entry.Value is String && de.Entry.Key is String)
                    {
                        result.Add((String)de.Entry.Key, (String)de.Entry.Value);
                    }
                }
            }

            return result;
        }

        public void Dispose()
        {
            if (this._typeDictionary != null && this._typeDictionary.Count > 0)
            {
                this._typeDictionary.Clear();
            }
        }
    }
}
