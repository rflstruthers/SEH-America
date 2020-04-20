using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using SehTest.Models;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Result = SehTest.Models.Result;
using Syncfusion.Presentation;

namespace SehTest.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Result(string title, string content)
        {

            WebClient webClient = new WebClient();

            string apiKey = "";
            string cx = "016005425349973865389:dsuv558pzqq";
            string query = title + " " + content;

            var request = WebRequest.Create("https://www.googleapis.com/customsearch/v1?key=" + apiKey + "&cx=" + cx + "&q=" + query + "&searchType=image");
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);
            string responseString = reader.ReadToEnd();
            dynamic jsonData = JsonConvert.DeserializeObject(responseString);

            var results = new List<Result>();
            foreach (var item in jsonData.items)
            {
                var images = item.image;
                results.Add(new Result
                {
                    Title = title,
                    Content = content,
                    Link = images.thumbnailLink
                }); ;
            }

            ViewBag.header = title;
            ViewBag.content = content;

            return View(results.ToList());

        }

        public IActionResult CreateSlide(string image, string header, string content)
        {
            WebClient webClient = new WebClient();
            webClient.DownloadFile(image, "image.jpg");

            //Create a new instance of PowerPoint Presentation file
            IPresentation pptxDoc = Presentation.Create();

            //Add a new slide to file and apply background color
            ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.TitleOnly);

            //Specify the fill type and fill color for the slide background 
            slide.Background.Fill.FillType = FillType.Solid;
            slide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(232, 241, 229);

            //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
            IShape titleShape = slide.Shapes[0] as IShape;
            titleShape.TextBody.AddParagraph(header).HorizontalAlignment = HorizontalAlignmentType.Center;

            //Add description content to the slide by adding a new TextBox
            IShape descriptionShape = slide.AddTextBox(53.22, 141.73, 874.19, 77.70);
            descriptionShape.TextBody.Text = content;

            //Gets a picture as stream.
            FileStream pictureStream = new FileStream("image.jpg", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the PowerPoint Presentation as stream
            FileStream outputStream = new FileStream("Sample.pptx", FileMode.Create);
            pptxDoc.Save(outputStream);

            //Release all resources from stream
            outputStream.Dispose();

            //Close the PowerPoint presentation
            pptxDoc.Close();

            var displaySlide = new Result
            {
                Title = header,
                Content = content,
                Link = image
            };

            return View(displaySlide);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
