namespace DocxReducer.Options
{
    public class ReducerOptions
    {
        public bool DeleteBookmarks { get; } = true;

        public bool CreateNewStyles { get; } = true;

        public ReducerOptions(bool deleteBookmarks, bool createNewStyles)
        {
            DeleteBookmarks = deleteBookmarks;
            CreateNewStyles = createNewStyles;
        }

        public ReducerOptions() { }
    }
}
