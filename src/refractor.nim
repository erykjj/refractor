const
  App = "refractor"
  Copyright = "© 2026 Eryk J."
  Version = "2.0.0"

#[  This code is licensed under the Infiniti Noncommercial License.
    You may use and modify this code for personal, non-commercial purposes only.
    Sharing, distribution, or commercial use is strictly prohibited.
    See LICENSE for full terms.                                              ]#

import
  std/[marshal, os, parseopt, strformat, strutils, tables, terminal, unicode, uri, xmlparser, xmltree],
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
    scriptures: seq[(string, string, string, string)]
    publications: seq[string]

  FocalizerPacket = object
    version: string
    languageId: int
    languageCode: string
    searchPath: string
    scriptureLangs: OrderedTable[string, (string, string, string)]
    publicationLangs: OrderedTable[string, (string, string, string)]

var
  lang = "en"
  inputFile = ""
  pkt: FocalizerPacket

proc init(languageCode, nameFormat: cstring): cstring {.cdecl, dynlib: libName, importc.}
proc extractAll(text: cstring): cstring {.cdecl, dynlib: libName, importc.}


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
  for word in reference.split(' '):
    let encodedWord = encodeUrl(word)
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
    let encoded = encodeForUrl(r.strip)
    stdout.styledWriteLine(fgGreen, &"{r.strip}", fgDefault, " --> ", fgBlue, url & encoded)
  echo ""

proc outputScriptureLinks(results: seq[(string, string, string, string)], url: string) =
  echo "\nOr use these individual links:\n"
  for item in results:
    let (_, alt, official, extra) = item
    var r = (alt & " " & extra).strip
    let encoded = encodeForUrl((official & " " & extra).strip)
    stdout.styledWriteLine(fgGreen, &"{r}", fgDefault, " --> ", fgBlue, url & encoded)
  echo ""

proc outputScriptures(results: seq[(string, string, string, string)]) =

  proc extractAlt(item: (string, string, string, string)): string =
    let (_, alt, _, extra) = item
    (alt & " " & extra).strip

  proc extractOfficial(item: (string, string, string, string)): string =
    let (_, _, official, extra) = item
    (official & " " & extra).strip

  let url = constructUrl()
  let searchChunks = createChunks(results, extractAlt)
  stdout.styledWriteLine("You can paste these into the search box on ", fgBlue, &"https://wol.jw.org/{lang}:")
  for chunk in searchChunks:
    styledEcho fgGreen, "\n" & chunk.join("; ")
  
  echo "\nOr use the link(s) to open wol.jw.org directly:"
  let urlChunks = createChunks(results, extractOfficial)
  for chunk in urlChunks:
    let combinedLinks = convertRefs(chunk)
    styledEcho fgBlue, "\n" & url & combinedLinks
  
  outputScriptureLinks(results, url)

proc outputPublications(results: seq[string]) =

  proc extract(item: string): string = item.strip

  let url = constructUrl()
  let chunks = createChunks(results, extract)
  outputChunks(chunks, url)
  outputPublicationLinks(results, url)

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

proc main(showScripts, showRefs: bool) =
  let source = readSource(inputFile)
  if source == "":
    styledEcho fgRed, "\n Error: Could not read input file or file is empty"
    return
  let serializedResults = extractAll(source.cstring)
  let results = to[ExtractionResults]($serializedResults)
  if showScripts:
    echo ""
    styledEcho fgYellow, $results.scriptures.len & " SCRIPTURE(S) FOUND\n"
    if results.scriptures.len > 0:
      outputScriptures(results.scriptures)
  if showRefs:
    echo ""
    styledEcho fgYellow, $results.publications.len & " PUBLICATION REFERENCE(S) FOUND\n"
    if results.publications.len > 0:
      outputPublications(results.publications)

when isMainModule:
  let
    appName = getAppFilename().split(sep)[^1]
    appHelp = unindent(&"""

      Usage: {appName} [-h | -v | -l] | [-r] [-s] [--full | --standard | --official] -c:code <infile>

      Options:
        -h, --help                      Show this help message and exit
        -v, --version                   Show the version and exit

        -c:<code>, --code=<code>        Language code or symbol (en by default)
        -l, --list                      List supported languages

        -r, --references                Output publication references
        -s, --scriptures                Output scriptures (if neither -r nor -s
                                          is provided, both shown)

      Scripture (book names) rewrite options:
        --full                          Use full name
        --standard                      Use standard name
        --official                      Use official name (default)

      <infile>                          File to process (docx or text)
      """, 5, " ")
  var
    showHelp = false
    showVersion = false
    showList = false
    isError = false
    showScripts = false
    showRefs = false
    nameFormat = "official"

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

  let serializedPacket = init(lang.cstring, nameFormat.cstring)
  pkt = to[FocalizerPacket]($serializedPacket) # language and book data
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
    main(showScripts, showRefs)
  finally:
    quit(0)
