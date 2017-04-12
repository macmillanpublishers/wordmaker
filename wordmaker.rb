# ---------------------- VARIABLES
unescapeargv = ARGV[0].chomp('"').reverse.chomp('"').reverse
inputfile = File.expand_path(unescapeargv)
inputfile = inputfile.split(Regexp.union(*[File::SEPARATOR, File::ALT_SEPARATOR].compact)).join(File::SEPARATOR)
working_dir = inputfile.split(Regexp.union(*[File::SEPARATOR, File::ALT_SEPARATOR].compact))[0...-2].join(File::SEPARATOR)
filetype = inputfile.split(".").pop
outputfile = File.join(working_dir, "output", "output.docx")
os = "windows"
resource_dir = "C:"
savehtmlasdocxps = ""
transformxmlpy = ""

# ---------------------- METHODS

def convertHMTLToDocxPSscript(psscript, filetype, htmlfile, docxfile)
  if filetype == "html"
    `PowerShell -NoProfile -ExecutionPolicy Bypass -Command "#{psscript} '#{inputfile}' '#{docxfile}'"`
  else
    logstring = 'input file is not html, skipping'
  end
rescue => logstring
ensure
  puts logstring
end

def runpython(py_script, dir, args)
  if filetype == "html"
    if os == "mac" or os == "unix"
      `python #{py_script} #{args}`
    elsif os == "windows"
      pythonpath = File.join(dir, "Python27", "python.exe")
      `#{pythonpath} #{py_script} #{args}`
    else
      File.open(Bkmkr::Paths.log_file, 'a+') do |f|
        f.puts "----- PYTHON ERROR"
        f.puts "ERROR: I can't seem to run python. Is it installed and part of your system PATH?"
        f.puts "ABORTING. All following processes will fail."
      end
      File.delete(Project.alert)
    end
  end
end

# ---------------------- PROCESSES

#convert .doc to .docx via powershell script, ignore html files
convertHMTLToDocxPSscript(savehtmlasdocxps, filetype, inputfile, outputfile)

runpython(transformxmlpy, resource_dir, "")