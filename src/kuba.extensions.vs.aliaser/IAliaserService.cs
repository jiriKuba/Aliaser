using System;

namespace kuba.extensions.vs.aliaser
{
    internal interface IAliaserService
    {
        String GetTransformedText(String selectedText, TransformationFlow flow);

        String GetTransformedText(String selectedText, TransformationFlow flow, out Boolean wasChangesMade);
    }
}