namespace DocxReducer.Options
{
    internal class ParagraphProcessorOptions
    {
        public bool DeleteBookmarks { get; set; }

        public bool CreateNewStyles { get; set; }

        public ParagraphProcessorOptions() { }

        public ParagraphProcessorOptions(ReducerOptions reducerOptions)
        {
            DeleteBookmarks = reducerOptions.DeleteBookmarks;
            CreateNewStyles = reducerOptions.CreateNewStyles;
        }

        public ParagraphProcessorOptions(bool deleteBookmarks, bool createNewStyles)
        {
            DeleteBookmarks = deleteBookmarks;
            CreateNewStyles = createNewStyles;
        }
    }
}
