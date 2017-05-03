require_relative '../bookmaker/core/utilities/mcmlln-tools.rb'

# ---------------------- VARIABLES
unescapeargv = ARGV[0].chomp('"').reverse.chomp('"').reverse
inputfile = File.expand_path(unescapeargv)
inputfile = inputfile.split(Regexp.union(*[File::SEPARATOR, File::ALT_SEPARATOR].compact)).join(File::SEPARATOR)
filename_split = inputfile.split(Regexp.union(*[File::SEPARATOR, File::ALT_SEPARATOR].compact)).pop
filename = inputfile.split(Regexp.union(*[File::SEPARATOR, File::ALT_SEPARATOR].compact)).pop.rpartition('.').first.gsub(/ /, "")
working_dir = inputfile.split(Regexp.union(*[File::SEPARATOR, File::ALT_SEPARATOR].compact))[0...-2].join(File::SEPARATOR)
final_dir = File.join(working_dir, "done", filename)
final_file = File.join(final_dir, filename_split)
filetype = inputfile.split(".").pop
outputfile = File.join(final_dir, "output.docx")
scriptsdir = File.join("S:", "resources", "bookmaker_scripts", "wordmaker")
worddir = File.join(final_dir, "word")
documentxml = File.join(worddir, "document.xml")
os = "windows"
resource_dir = "C:"
newzipdir = File.join(final_dir, "newzip")
htmlconversionsjs = File.join(scriptsdir, "HTMLconversions.js")
savehtmlasdocxps = File.join(scriptsdir, "saveHTMLasDOCX.ps1")
unzipdocxpy = File.join(scriptsdir, "unzipDOCX.py")
renamestylespy = File.join(scriptsdir, "renameStyles.py")

# ---------------------- METHODS

def makeDir(path)
  unless Dir.exist?(path)
    Mcmlln::Tools.makeDir(path)
  else
    logstring = 'n-a'
  end
rescue => logstring
ensure
  puts logstring
end

def moveFile(file, dest)
  Mcmlln::Tools.moveFile(file, dest)
rescue => logstring
ensure
  puts logstring
end

def convertHMTLToDocxPSscript(psscript, filetype, file, docxfile)
  if filetype == "html"
    `PowerShell -NoProfile -ExecutionPolicy Bypass -Command "#{psscript} '#{file}' '#{docxfile}'"`
  else
    logstring = 'input file is not html, skipping'
  end
rescue => logstring
ensure
  puts logstring
end

def localRunNode(jsfile, args, os, resource_dir)
  if os == "mac" or os == "unix"
    `node #{jsfile} #{args}`
  elsif os == "windows"
    nodepath = File.join(resource_dir, "nodejs", "node.exe")
    puts "running node"
    nodecmd = `#{nodepath} #{jsfile} #{args}`
    puts nodecmd
  else
    puts "Can't run node!"
  end
rescue => logstring
ensure
  puts logstring
end

def runpython(py_script, dir, os, args)
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

# ---------------------- PROCESSES

puts "----------NEW RUN-----------"

# create the folder for doing everything
makeDir(final_dir)

# move input file to final dir
moveFile(inputfile, final_dir)

# prep the html file for conversion
puts final_file
localRunNode(htmlconversionsjs, final_file, os, resource_dir)

# convert .html to .docx via powershell script
convertHMTLToDocxPSscript(savehtmlasdocxps, filetype, final_file, outputfile)

# create the folder for doing everything
makeDir(newzipdir)

# unzip converted docx
runpython(unzipdocxpy, resource_dir, os, "#{outputfile} #{newzipdir}")

# rename styles in document.xml
#runpython(renamestylespy, resource_dir, os, "'#{documentxml}' '#{newzipdir}'")

# copy styles file into new zip
#File.copy("styles.xml", worddir)