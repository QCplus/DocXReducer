namespace DocxReducer.Options
{
    public class ReducerOptions
    {
        public bool DeleteBookmarks { get; set; }

        public bool CreateNewStyles { get; set; }

        public ReducerOptions(bool deleteBookmarks, bool createNewStyles)
        {
            DeleteBookmarks = deleteBookmarks;
            CreateNewStyles = createNewStyles;
        }

        public ReducerOptions()
        {
            
        }
    }
}
