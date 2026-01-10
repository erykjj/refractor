const
  App = "refractor"
  Copyright = "© 2025 Eryk J."
  Version = "1.2.1"

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
    scriptures: seq[(string, string, string)]
    publications: seq[(string, string, string)]
    allInOrder: seq[(string, string, string)]

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

proc init(languageCode: cstring): cstring {.cdecl, dynlib: libName, importc.}
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

proc output(results: seq[(string, string, string)]) = 
  let url = constructUrl()
  var chunks: seq[seq[string]] = @[]
  var currentChunk: seq[string] = @[]
  var currentLength = 0
  for (source, alt, extra) in results:
    var r = ""
    if alt.len > 0:
      r = (alt & " " & extra).strip
    else:
      r = (source & " " & extra).strip
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
  stdout.styledWriteLine("You can paste these into the search box on ", fgBlue, &"https://wol.jw.org/{lang}:")
  for i, chunk in chunks:
    styledEcho fgGreen, "\n" & chunk.join("; ")
  echo "\nOr use the link(s) to open wol.jw.org directly:"
  for i, chunk in chunks:
    let combinedLinks = convertRefs(chunk)
    styledEcho fgBlue, "\n" & url & combinedLinks
  echo "\nOr use these individual links:\n"
  for i, chunk in chunks:
    for r in chunk:
      let encoded = encodeForUrl(r)
      stdout.styledWriteLine(fgGreen, &"{r}", fgDefault, " --> ", fgBlue, url & encoded)
    echo ""

proc languageList(list: OrderedTable[string, (string, string, string)]) =
  var t = tabulator.newTable()
  t.addColumn(width=24)
  t.addColumn(width=5)
  t.addColumn(width=5)
  t.addColumn()
  for code, names in list:
    var (symbol, name, vernacular) = names
    t.addRow(@[" " & name, &"\e[32m{code}\e[0m", &"\e[32m{symbol}\e[0m", vernacular])
  t.renderTable(separator=false)

proc main(showScripts, showRefs: bool) =
  let source = readSource(inputFile)
  if source == "":
    styledEcho fgRed, "\n Error: Could not read input file or file is empty"
    return
  let serializedResults = extractAll(source.cstring)
  let results = to[ExtractionResults]($serializedResults)
  if not showRefs and not showScripts:
    styledEcho fgYellow, $results.allInOrder.len & " reference(s) found"
    if results.allInOrder.len > 0:
      output(results.allInOrder)
  else:
    if showScripts:
      styledEcho fgYellow, $results.scriptures.len & " scripture(s) found\n"
      if results.scriptures.len > 0:
        output(results.scriptures)
    if showRefs:
      styledEcho fgYellow, $results.publications.len & " publication reference(s) found\n"
      if results.publications.len > 0:
        output(results.publications)

when isMainModule:
  let
    appName = getAppFilename().split(sep)[^1]
    appHelp = unindent(&"""

      Usage: {appName} [-h | -v | -l] | [-r] [-s] -c:code <infile>

      Options:
        -h, --help                      Show this help message and exit
        -v, --version                   Show the version and exit

        -c:<code>, --code=<code>        Language code or symbol (en by default)
        -l, --list                      List supported languages

        -r, --references                Output publication references
        -s, --scriptures                Output scriptures (if neither -r or
                                          -s is provided, both are enabled)

      <infile>                          File to process (docx or text)
      """, 5, " ")
  var
    showHelp = false
    showVersion = false
    showList = false
    isError = false
    showScripts = false
    showRefs = false

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
      else:
        styledEcho fgRed, &"\n Error: Unknown option '{key}'"
        isError = true
    of cmdEnd:
      discard

  let serializedPacket = init(lang.cstring)
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

  if showList:
    stdout.styledWriteLine(fgBlue, &"\n Supported scripture languages ({$len(pkt.scriptureLangs)}):")
    languageList(pkt.scriptureLangs)
    stdout.styledWriteLine(fgBlue, &"\n Supported publication languages ({$len(pkt.publicationLangs)}):")
    languageList(pkt.publicationLangs)
    quit(0)

  if inputFile == "":
    styledEcho fgRed, " Error: provide an input/source file"
    isError = true

  if isError:
    echo &"\n See '{appName} -h' for help.\n"
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
