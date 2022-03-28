-- Startup script
-- Changes will take effect once Notepad++ is restarted

npp.AddShortcut("Multiple Cursor - Selection Add Next", "", function()
  -- From SciTEBase.cxx
  local flags = SCFIND_WHOLEWORD -- can use 0

  editor:SetTargetRange(0, editor.TextLength)
  editor.SearchFlags = flags

  -- From Editor.cxx
  if editor.SelectionEmpty or not editor.MultipleSelection then
    local startWord = editor:WordStartPosition(editor.CurrentPos, true)
    local endWord   = editor:WordEndPosition(startWord, true)

    editor:SetSelection(startWord, endWord)
  else
    local i = editor.MainSelection
    local s = editor:textrange(editor.SelectionNStart[i], editor.SelectionNEnd[i])
    local searchRanges = {{editor.SelectionNEnd[i], editor.TargetEnd}, {editor.TargetStart, editor.SelectionNStart[i]}}

    for _, range in pairs(searchRanges) do
      editor:SetTargetRange(range[1], range[2])

      if editor:SearchInTarget(s) ~= -1 then
        editor:AddSelection(editor.TargetStart, editor.TargetEnd)

        -- To scroll main selection in sight
        editor:ScrollRange(editor.TargetStart, editor.TargetEnd)

        break
      end
    end
  end

  -- To turn on Notepad++ multi select markers
  editor:LineScroll(0, 1)
  editor:LineScroll(0, -1)
end)


npp.AddShortcut("Multiple Cursor - Selection Add All", "", function()
  local flags     = SCFIND_WHOLEWORD -- can use 0
  local startWord = -1
  local endWord   = -1
  local s         = ""

  editor.SearchFlags = flags

  if editor.SelectionEmpty or not editor.MultipleSelection then
    startWord = editor:WordStartPosition(editor.CurrentPos, true)
    endWord   = editor:WordEndPosition(startWord, true)

    editor:SetSelection(startWord, endWord)
  else
    local i   = editor.MainSelection

    startWord = editor.SelectionNStart[i]
    endWord   = editor.SelectionNEnd[i]
  end

  s = editor:textrange(startWord, endWord)

  while true do
    editor:SetTargetRange(0, editor.TextLength)

    local i            = editor.MainSelection
    local searchRanges = {{editor.SelectionNEnd[i], editor.TargetEnd}, {editor.TargetStart, editor.SelectionNStart[i]}}
    local itemFound    = false

    for _, range in pairs(searchRanges) do
      editor:SetTargetRange(range[1], range[2])

      if editor:SearchInTarget(s) ~= -1 then
        editor:AddSelection(editor.TargetStart, editor.TargetEnd)
        itemFound = true
        break
      end
    end

    if editor.TargetStart == startWord and
       editor.TargetEnd   == endWord   or
       not itemFound                   then
      break
    end
  end

  -- To turn on Notepad++ multi select markers
  editor:LineScroll(0, 1)
  editor:LineScroll(0, -1)
end)
