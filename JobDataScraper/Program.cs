using JobDataScraper.Services;

namespace JobDataScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            //new InappService().Main();
            new InappJobAtlasService().Main();
        }
    }
}
