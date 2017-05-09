var fs = require('fs');
var cheerio = require('cheerio');
var file = process.argv[2];

fs.readFile(file, function editContent (err, contents) {
  $ = cheerio.load(contents, {
          xmlMode: true
        });

// Remove extra classes for footnote refs
  $(".FootnoteRef").removeClass().addClass("FootnoteRef");

// convert footnote spans to paragraphs
  $("div.footnote span.FootnoteText").each(function () {
    var myHtml = $(this).html();
    $(this).replaceWith(function(){
        return $("<p/>").html(myHtml).addClass("FootnoteText");
    });
  });

// delete any empty spans
  $("span:empty").remove();

  var output = $.html();
    fs.writeFile(file, output, function(err) {
      if(err) {
          return console.log(err);
      }

      console.log("Content has been updated!");
  });
});