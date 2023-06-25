namespace DocxReducer.Options
{
    public class ReducerOptions
    {
        public bool DeleteBookmarks { get; }

        public bool CreateNewStyles { get; }

        public ReducerOptions(bool deleteBookmarks, bool createNewStyles)
        {
            DeleteBookmarks = deleteBookmarks;
            CreateNewStyles = createNewStyles;
        }
    }
}
