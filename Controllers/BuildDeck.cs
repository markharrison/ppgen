using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Mvc;
using OpenXmlPowerTools;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace PowerPointGeneration.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class BuildDeck : ControllerBase
    {
        public static IWebHostEnvironment? _environment;
        public static IConfiguration? _config;
        public BuildDeck(IWebHostEnvironment environment, IConfiguration config)
        {
            _environment = environment;
            _config = config;
        }


        [HttpPost]
        [ProducesResponseType(StatusCodes.Status201Created)]
        //public IActionResult Post([FromForm] BusinessValueModel jsonpayload)
        public async Task<IActionResult> Post([FromBody] BusinessValueModel bvmodel)
        {
            await Task.Run(() => { });
            try
            {
                Guid newGuid = Guid.NewGuid();

                string resourcedirectory = _environment.ContentRootPath + "/Resources/";
                PmlDocument DeckA = new PmlDocument(resourcedirectory + "DeckA.pptx");
                PmlDocument DeckB = new PmlDocument(resourcedirectory + "DeckB.pptx");

                List<SlideSource> MainDeck = new List<SlideSource>(); //Create List of Slide Source to Merge Slides
                MainDeck.Add(new SlideSource(UpdateSlide(DeckA, bvmodel), true));
                MainDeck.Add(new SlideSource(DeckB, true));
                PmlDocument FinalDeck = PresentationBuilder.BuildPresentation(MainDeck);

                var vFilename = "Report-" + newGuid + ".pptx";

                FinalDeck.SaveAs("wwwroot/"+ vFilename);

                return Created("PPT", _config.GetValue<string>("ServerURL") + vFilename);

            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }

        }

        private PmlDocument UpdateSlide(PmlDocument deckA, BusinessValueModel? bvmodel)
        {
            PmlDocument? modifiedMainPresentation = null;

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(deckA))
            {
                using (PresentationDocument document = streamDoc.GetPresentationDocument())
                {
                    var pXDoc = document.PresentationPart.GetXDocument();
                    foreach (var slideId in pXDoc.Root.Elements(P.sldIdLst).Elements(P.sldId))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        var slidePart = document.PresentationPart.GetPartById(slideRelId);
                        var slideXDoc = slidePart.GetXDocument();
                        var paragraphs = slideXDoc.Descendants(A.p).ToList();

                        OpenXmlRegex.Replace(paragraphs, new Regex("<customer>"), bvmodel.Customer, null);
                        OpenXmlRegex.Replace(paragraphs, new Regex("<title>"), bvmodel.Title, null);
                        slidePart.PutXDocument();
                    }
                }
                modifiedMainPresentation = streamDoc.GetModifiedPmlDocument();
            }
            return modifiedMainPresentation;
        }

        public class BusinessValueModel
        {
            public string? Customer { get; set; }
            public string? Title { get; set; }
        }
    }
}
