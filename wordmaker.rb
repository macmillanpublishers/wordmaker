# ---------------------- VARIABLES
unescapeargv = ARGV[0].chomp('"').reverse.chomp('"').reverse
inputfile = File.expand_path(unescapeargv)
inputfile = inputfile.split(Regexp.union(*[File::SEPARATOR, File::ALT_SEPARATOR].compact)).join(File::SEPARATOR)
working_dir = inputfile.split(Regexp.union(*[File::SEPARATOR, File::ALT_SEPARATOR].compact))[0...-2].join(File::SEPARATOR)
filetype = inputfile.split(".").pop
outputfile = File.join(working_dir, "output", "output.docx")

# ---------------------- METHODS

def convertHMTLToDocxPSscript(filetype, htmlfile, docxfile)
  if filetype == "html"
    `PowerShell -NoProfile -ExecutionPolicy Bypass -Command "#{htmltodocxps} '#{inputfile}' '#{docxfile}'"`
  else
    logstring = 'input file is not html, skipping'
  end
rescue => logstring
ensure
  puts logstring
end

# ---------------------- PROCESSES

#convert .doc to .docx via powershell script, ignore html files
convertHMTLToDocxPSscript(filetype, inputfile, outputfile)