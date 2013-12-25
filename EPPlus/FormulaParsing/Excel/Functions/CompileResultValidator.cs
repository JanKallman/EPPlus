namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public abstract class CompileResultValidator
    {
        public abstract void Validate(object obj);

        private static CompileResultValidator _empty;
        public static CompileResultValidator Empty
        {
            get { return _empty ?? (_empty = new EmptyCompileResultValidator()); }
        }
    }

    internal class EmptyCompileResultValidator : CompileResultValidator
    {
        public override void Validate(object obj)
        {
            // empty validator - do nothing
        }
    }
}
