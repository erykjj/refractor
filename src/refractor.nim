const
  App = "refractor"
  Copyright = "© 2026 Eryk J."
  Version = "3.1.0"

#[  This code is licensed under the Infiniti Noncommercial License.
    You may use and modify this code for personal, non-commercial purposes only.
    Sharing, distribution, or commercial use is strictly prohibited.
    See LICENSE for full terms.                                              ]#

import
  std/[algorithm, marshal, options, os, parseopt, strformat, strutils, tables, terminal, unicode, uri, xmlparser, xmltree],
  tabulator,
  zippy/ziparchives

when defined(windows):
  const
    libName = "focalizer.dll"
    sep = r"\"
elif defined(macosx):
  const
    libName = "libfocalizer.dylib"
    sep = "/"
else: # linux
  const
    libName = "./libfocalizer.so"
    sep = "/"

type
  ExtractionResults = object
    scriptures: seq[(string, string)]
    publications: seq[string]

  FocalizerPacket = object
    version: string
    languageId: int
    languageCode: string
    searchPath: string
    scriptureLangs: OrderedTable[string, (string, string, string)]
    publicationLangs: OrderedTable[string, (string, string, string)]

  Config = object
    languageCode: Option[string]
    nameFormat: Option[string]
    showReferences: Option[bool]
    showScriptures: Option[bool]

var
  lang = "en"
  inputFile = ""
  pkt: FocalizerPacket


proc focus(languageCode, nameFormat: cstring): cstring {.cdecl, dynlib: libName, importc.}
proc extractAll(text: cstring, sortedOutput: bool): cstring {.cdecl, dynlib: libName, importc.}


proc docxOpen(docxFile: string): string =
  let reader = openZipArchive(docxFile)
  try:
    result = reader.extractFile("word/document.xml")
  finally:
    reader.close()

proc docxRead(docxFile: string): string =
  let xml = parseXml(docxOpen(docxFile))
  var text: string
  for item in xml.findAll("w:t"):
    text.add(item.innerText & " ")
  result = text.replace("  ", " ").replace("  ", " ")

proc readSource(filePath: string): string =
  if not fileExists(filePath):
    return ""
  try:
    if filePath.endsWith(".docx"):
      result = docxRead(filePath)
    else:
      result = readFile(filePath)
  except IOError, OSError:
    result = ""


proc constructUrl(): string =
  if pkt.searchPath == "":
    return ""
  result = "https://wol.jw.org/" & lang & "/wol/l" & pkt.searchPath & "?q="

proc encodeForUrl(reference: string): string =
  var parts: seq[string] = @[]
  for word in reference.strip().split(' '):
    let encodedWord = encodeUrl(word.replace("‑", "-"))
    parts.add(encodedWord)
  result = parts.join("+")

proc convertRefs(refList: seq[string]): string =
  var encoded: seq[string] = @[]
  for reference in refList:
    encoded.add(encodeForUrl(reference))
  result = encoded.join(";")

proc createChunks[T](items: seq[T], extractor: proc(item: T): string): seq[seq[string]] =
  var chunks: seq[seq[string]] = @[]
  var currentChunk: seq[string] = @[]
  var currentLength = 0
  for item in items:
    let r = extractor(item)
    let rLen = r.len
    let additionalLength = if currentLength == 0: rLen else: rLen + 2
    if currentLength + additionalLength > 255 and currentChunk.len > 0:
      chunks.add(currentChunk)
      currentChunk = @[r]
      currentLength = rLen
    else:
      currentChunk.add(r)
      currentLength += additionalLength
  if currentChunk.len > 0:
    chunks.add(currentChunk)
  return chunks

proc outputChunks(chunks: seq[seq[string]], url: string) =
  stdout.styledWriteLine("You can paste these into the search box on ", fgBlue, &"https://wol.jw.org/{lang}:")
  for chunk in chunks:
    styledEcho fgGreen, "\n" & chunk.join("; ")
  echo "\nOr use the link(s) to open wol.jw.org directly:"
  for chunk in chunks:
    let combinedLinks = convertRefs(chunk)
    styledEcho fgBlue, "\n" & url & combinedLinks

proc outputPublicationLinks(items: seq[string], url: string) =
  echo "\nOr use these individual links:\n"
  for r in items:
    let encoded = encodeForUrl(r.strip())
    stdout.styledWriteLine(fgGreen, &"{r.strip()}", fgDefault, " --> ", fgBlue, url & encoded)
  echo ""

proc outputScriptureLinks(results: seq[(string, string)], url: string) =
  echo "\nOr use these individual links:\n"
  for item in results:
    let (_, chosen) = item
    let encoded = encodeForUrl(chosen)
    stdout.styledWriteLine(fgGreen, &"{chosen}", fgDefault, " --> ", fgBlue, url & encoded)
  echo ""

proc extractOfficial(item: (string, string)): string =
  let (official, _) = item
  official

proc extract(item: string): string =
  item.strip

proc outputScriptures(results: seq[(string, string)]) =
  let url = constructUrl()
  let searchChunks = createChunks(results, extractOfficial)
  stdout.styledWriteLine("You can paste these into the search box on ", fgBlue, &"https://wol.jw.org/{lang}:")
  for chunk in searchChunks:
    styledEcho fgGreen, "\n" & (chunk.join(";")).replace(" ", "")

  echo "\nOr use the link(s) to open wol.jw.org directly:"
  let urlChunks = createChunks(results, extractOfficial)
  for chunk in urlChunks:
    let combinedLinks = convertRefs(chunk)
    styledEcho fgBlue, "\n" & url & combinedLinks
  outputScriptureLinks(results, url)

proc outputPublications(results: seq[string]) =
  let url = constructUrl()
  let chunks = createChunks(results, extract)
  outputChunks(chunks, url)
  outputPublicationLinks(results, url)


proc generateHtmlOutput(results: ExtractionResults, showScripts, showRefs: bool) =
  let htmlPath = getAppDir() / "refractor_output.html"
  let url = constructUrl()
  var html = unindent("""
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <title>refractor output</title>
        <style>
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 2rem; line-height: 1.6; }
          h1 { color: #333; border-bottom: 2px solid #ddd; padding-bottom: 0.5rem; }
          h2 { color: #666; margin-top: 2rem; }
          hr { border: none; border-top: 1px solid #ddd; margin: 2rem 0; }
          .instruction { color: #666; margin: 1rem 0; }
          .chunk { display: block; margin: 0.75rem 0; padding: 0.75rem; background: #f5f5f5; border-left: 4px solid #496DA7; font-family: monospace; }
          a { color: #496DA7; text-decoration: none; }
          a:hover { text-decoration: underline; }
          .individual-links { margin-top: 1rem; }
          .individual-links a { display: inline-block; margin: 0.25rem 0.5rem; padding: 0.25rem 0.5rem; background: #f0f0f0; color: #496DA7; text-decoration: none; border-radius: 3px; }
          .individual-links a:hover { background: #e0e0e0; text-decoration: underline; }
        </style>
      </head>
      <body>
  """, 4, " ")
  html.add("    <h1>refractor output</h1>\n")

  # Scriptures section
  if showScripts and results.scriptures.len > 0:
    html.add(&"    <h2>Scripture references ({results.scriptures.len})</h2>\n")
    html.add(&"    <div class='instruction'>Grouped search on <a href='https://wol.jw.org/{lang}/' target='_blank'>wol.jw.org</a>:</div>\n")
    let urlChunks = createChunks(results.scriptures, extractOfficial)
    for chunk in urlChunks:
      let combinedLinks = convertRefs(chunk)
      let fullUrl = url & combinedLinks
      html.add(&"    <div class='chunk'><a href='{fullUrl}' target='_blank'>{chunk.join(\"; \")}</a></div>\n")
    html.add("    <div class='individual-links'>\n")
    html.add("      <div class='instruction'>Individual links:</div>\n")
    for item in results.scriptures:
      let (official, chosen) = item
      let encoded = encodeForUrl(official)
      let fullUrl = url & encoded
      html.add(&"      <a href='{fullUrl}' target='_blank'>{chosen}</a>\n")
    html.add("    </div>\n")

  # Publications section
  if showRefs and results.publications.len > 0:
    html.add("    <hr />\n")
    html.add(&"    <h2>Publication references ({results.publications.len})</h2>\n")
    html.add(&"    <div class='instruction'>Grouped search on <a href='https://wol.jw.org/{lang}/' target='_blank'>wol.jw.org</a>:</div>\n")
    let chunks = createChunks(results.publications, extract)
    for chunk in chunks:
      let combinedLinks = convertRefs(chunk)
      let fullUrl = url & combinedLinks
      html.add(&"    <div class='chunk'><a href='{fullUrl}' target='_blank'>{chunk.join(\"; \")}</a></div>\n")
    html.add("    <div class='individual-links'>\n")
    html.add("      <div class='instruction'>Individual links:</div>\n")
    for item in results.publications:
      let encoded = encodeForUrl(item)
      let fullUrl = url & encoded
      html.add(&"      <a href='{fullUrl}' target='_blank'>{item}</a>\n")
    html.add("    </div>\n")
  html.add("  </body>\n</html>\n")
  try:
    writeFile(htmlPath, html)
  except:
    styledEcho fgYellow, &"\n Warning: Could not write HTML output to {htmlPath}\n"


proc languageList(list: OrderedTable[string, (string, string, string)]) =
  var t = tabulator.newTable()
  t.addColumn(width=22)
  t.addColumn(width=22)
  t.addColumn(width=3)
  t.addColumn(width=3)
  for code, names in list:
    var (symbol, name, vernacular) = names
    t.addRow(@[" " & name, vernacular, &"\e[32m{code}\e[0m", &"\e[32m{symbol}\e[0m"])
  t.renderTable(separator=false)

proc createDefaultConfig(configPath: string) =
  let defaultConfig = unindent(&"""
    # refractor Configuration File
    # Edit this file to set your preferred defaults

    # Language code (run with -l to see supported languages)
    # Examples: 'en' or 'E', 'es' or 'S' for Spanish, etc.
    code=en

    # Book name format:
    #   full      - Full book name (e.g., "Genesis")
    #   standard  - Standard abbreviation (e.g., "Gen.")
    #   official  - Official abbreviation (e.g., "Ge")
    format=official

    # Output preferences (default: both shown)
    # Uncomment to enable only one output type
    # references=true
    # scriptures=true
    """, 4, " ")
  try:
    writeFile(configPath, defaultConfig)
  except:
    discard

proc loadConfig(): Config =
  let configPath = getAppDir() / "refractor.conf"
  if not fileExists(configPath):
    createDefaultConfig(configPath)
  if getFileSize(configPath) == 0:
    styledEcho fgYellow, &"\n Warning: Could not parse config file '{configPath}'\n Using defaults\n"
    return
  result = Config(
    languageCode: none(string), 
    nameFormat: none(string),
    showReferences: none(bool),
    showScriptures: none(bool)
  )
  try:
    for line in lines(configPath):
      let trimmed = line.strip()
      if trimmed.len == 0 or trimmed.startsWith('#'):
        continue
      let parts = trimmed.split('=', 1)
      if parts.len == 2:
        let key = parts[0].strip().toLowerAscii()
        let value = parts[1].strip()
        case key
        of "code":
          if value.len > 0:
            result.languageCode = some(value)
        of "format":
          let format = value.toLowerAscii()
          if format in ["full", "standard", "official"]:
            result.nameFormat = some(format)
        of "references":
          if value.toLowerAscii() == "true":
            result.showReferences = some(true)
        of "scriptures":
          if value.toLowerAscii() == "true":
            result.showScriptures = some(true)
  except:
    styledEcho fgYellow, &"\n Warning: Could not parse config file '{configPath}'\n Using defaults\n"
    return

proc main(showScripts, showRefs, sortedOutput: bool): ExtractionResults =
  let source = readSource(inputFile)
  if source == "":
    styledEcho fgRed, "\n Error: Could not read input file or file is empty"
    return
  let serializedResults = extractAll(source.cstring, sortedOutput)
  var results = to[ExtractionResults]($serializedResults)
  if showScripts:
    echo ""
    styledEcho fgYellow, $results.scriptures.len & " SCRIPTURE(S) FOUND\n"
    if results.scriptures.len > 0:
      outputScriptures(results.scriptures)
  if showRefs:
    echo ""
    styledEcho fgYellow, $results.publications.len & " PUBLICATION REFERENCE(S) FOUND\n"
    if results.publications.len > 0:
      if sortedOutput:
        var sorted = results.publications.sorted(cmp)
        var deduped: seq[string] = @[]
        for i in 0..<sorted.len:
          if i == 0 or sorted[i] != sorted[i-1]:
            deduped.add(sorted[i])
        results.publications = deduped
      outputPublications(results.publications)
  return results


when isMainModule:
  let
    appName = getAppFilename().split(sep)[^1]
    appHelp = unindent(&"""

      Usage: {appName} [-h | -v | -l] | [-r] [-s] [--sorted] [--full | --standard | --official] -c:code <infile>

      Options:
        -h, --help                      Show this help message and exit
        -v, --version                   Show the version and exit

        -c:<code>, --code=<code>        Language code or symbol (en by default)
        -l, --list                      List supported languages

        -r, --references                Output publication references
        -s, --scriptures                Output scriptures (if neither -r nor -s
                                          is provided, both shown)
        --sorted                        Sort output (scriptures by book order,
                                          publications alphabetically)

      Scripture (book names) rewrite options:
        --full                          Use full name
        --standard                      Use standard name
        --official                      Use official name (default)

      <infile>                          File to process (docx or text)
      """, 5, " ")

  let config = loadConfig()
  var
    showHelp = false
    showVersion = false
    showList = false
    isError = false
    showScripts = config.showScriptures.get(false)
    showRefs = config.showReferences.get(false)
    nameFormat = config.nameFormat.get("official")
    sortedOutput = false
  lang = config.languageCode.get("en")

  for kind, key, val in getOpt():
    case kind
    of cmdArgument:
      inputFile = key
    of cmdLongOption, cmdShortOption:
      case key
      of "code", "c":
        lang = val
      of "scriptures", "s":
        showScripts = true
      of "references", "r":
        showRefs = true
      of "help", "h":
        showHelp = true
      of "version", "v":
        showVersion = true
      of "list", "l":
        showList = true
      of "sorted":
        sortedOutput = true
      of "standard":
        nameFormat = "standard"
      of "full":
        nameFormat = "full"
      of "official":
        nameFormat = "official"
      else:
        isError = true
    of cmdEnd:
      discard
  let serializedPacket = focus(lang.cstring, nameFormat.cstring)
  pkt = to[FocalizerPacket]($serializedPacket)
  if pkt.languageCode == "":
    styledEcho fgRed, &"\n Error: language code '{lang}' not available"
    echo &"\n See '{appName} -l' for list of available languages.\n"
    quit(0)
  lang = pkt.languageCode

  if showHelp:
    stdout.styledWriteLine(fgBlue, &"\n {App} ", fgGreen, "REFERENCE EXTRACTOR ", fgDefault, "for publications of Jehovah's Witnesses")
    echo appHelp
    quit(0)

  if showVersion:
    styledEcho fgBlue, &"\n {App} v{Version}"
    styledEcho fgYellow, " " & pkt.version
    echo &"  {Copyright}\n"
    quit(0)

  if inputFile == "":
    styledEcho fgRed, " Error: provide an input/source file"
    isError = true

  if isError:
    echo &"\n See '{appName} -h' for help.\n"
    quit(0)

  if showList:
    stdout.styledWriteLine(fgBlue, &"\n Supported scripture languages ({$len(pkt.scriptureLangs)}):")
    languageList(pkt.scriptureLangs)
    stdout.styledWriteLine(fgBlue, &"\n Supported publication languages ({$len(pkt.publicationLangs)}):")
    languageList(pkt.publicationLangs)
    quit(0)

  if showScripts or showRefs:
    if showScripts and lang notin pkt.scriptureLangs:
      styledEcho fgRed, "\n Error: language code not available for parsing scripture references"
      echo &"\n See '{appName} -l' for list of available languages.\n"
      showScripts = false
    if showRefs and lang notin pkt.publicationLangs:
      styledEcho fgRed, "\n Error: language code not available for parsing publication references"
      echo &"\n See '{appName} -l' for list of available languages.\n"
      showRefs = false
    if not showScripts and not showRefs:
      quit(0)
  else:
    showScripts = lang in pkt.scriptureLangs
    showRefs = lang in pkt.publicationLangs
    if not showScripts and not showRefs:
      styledEcho fgRed, "\n Error: language code not available"
      echo &"\n See '{appName} -l' for list of available languages.\n"
      quit(0)

  try:
    let results = main(showScripts, showRefs, sortedOutput)
    generateHtmlOutput(results, showScripts, showRefs)
  finally:
    quit(0)
