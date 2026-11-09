-- VisualTeX live paragraphs. A LuaLaTeX run that has read the document's preamble (and \begin{document}, so
-- its labels are known) waits here for paragraphs on stdin. Each is typeset as the document typesets it (the
-- line width and indentation of the place it stands) and its lines come back: where every glyph and rule is,
-- from the top left of the paragraph, in sp. tex-bridge draws them over the page until TeX has made it anew.
--
-- A request:  "P <id> <hsize> <indent> <leftskip> <rightskip> <bytes>\n" then <bytes> of LaTeX
--   (indent: -1 the document's \parindent, 0 no indent, -2 the text says, else the indent in sp)
-- "R <file>\n": read the labels of the document's last run again (its .aux)
-- "Q\n": the run ends.
-- An answer: "\n@@VTXLIVE@@ <json>\n" on stdout.
local GLYPH, HLIST, VLIST, RULE, GLUE, KERN, DISC, MATH, INS, MARK, ADJUST, WHATSIT =
  node.id("glyph"), node.id("hlist"), node.id("vlist"), node.id("rule"), node.id("glue"), node.id("kern"),
  node.id("disc"), node.id("math"), node.id("ins"), node.id("mark"), node.id("adjust"), node.id("whatsit")
local RUNNING = -1073741824
local fonts_sent = {}

local function esc(s)
  return (s:gsub('[%c"\\]', function(c) return string.format("\\u%04x", c:byte()) end))
end

-- The map file's entries (pdftex.map: TFM name -> Type1 file, encoding, effects), read once when a TFM font
-- needs one.
local map_entries = nil
local function map_entry(name)
  if not map_entries then
    map_entries = {}
    local file = kpse.find_file("pdftex.map", "map")
    local h = file and io.open(file, "r")
    if h then
      for line in h:lines() do
        local first = line:sub(1, 1)
        if first ~= "%" and first ~= "#" and first ~= "" then
          local tfm = line:match("^(%S+)")
          if tfm and not map_entries[tfm] then map_entries[tfm] = line end
        end
      end
      h:close()
    end
  end
  local line = map_entries[name]
  if not line then return nil end
  local entry = { slant = 0, extend = 1 }
  local special = line:match('"([^"]*)"')
  if special then
    entry.slant = tonumber(special:match("([%-%d.]+)%s+SlantFont")) or 0
    entry.extend = tonumber(special:match("([%-%d.]+)%s+ExtendFont")) or 1
  end
  for file in line:gmatch("<[<%[]?([^%s<]+)") do
    local lower = file:lower()
    if lower:match("%.enc$") then entry.enc = kpse.find_file(file, "enc files") or ""
    elseif lower:match("%.pf[ab]$") then entry.font = kpse.find_file(file, "type1 fonts") or ""
    elseif lower:match("%.[ot]tf$") then entry.font = kpse.find_file(file, "opentype fonts") or kpse.find_file(file, "truetype fonts") or ""
    end
  end
  return entry
end

-- what tex-bridge needs to draw a font's glyphs: an OpenType/TrueType file and glyph indices ("otf"), or a
-- Type1 file, its encoding and the TFM's slots ("type1"); "none": not drawable here (a bitmap font)
local function font_entry(id)
  local f = font.getfont(id) or {}
  local file = f.filename or ""
  local lower = file:lower()
  local size = f.size or 0
  if lower:match("%.[ot]tf$") or lower:match("%.tt[cf]$") or lower:match("%.otc$") then
    return string.format('{"id":%d,"kind":"otf","file":"%s","sub":%d,"size":%d,"slant":%d,"extend":%d,"name":"%s"}',
      id, esc(file), f.subfont and math.max(f.subfont - 1, 0) or 0, size, f.slant or 0, f.extend or 1000, esc(f.name or ""))
  end
  local entry = map_entry(f.name or "")
  if entry and entry.font and entry.font ~= "" then
    local kind = entry.font:lower():match("%.pf[ab]$") and "type1" or "otf"
    return string.format('{"id":%d,"kind":"%s","file":"%s","enc":"%s","sub":0,"size":%d,"slant":%d,"extend":%d,"name":"%s"}',
      id, kind, esc(entry.font), esc(entry.enc or ""), size, math.floor(entry.slant * 1000 + 0.5),
      math.floor(entry.extend * 1000 + 0.5), esc(f.name or ""))
  end
  return string.format('{"id":%d,"kind":"none","size":%d,"name":"%s"}', id, size, esc(f.name or ""))
end

local function note_font(id, fonts)
  if not fonts_sent[id] then
    fonts_sent[id] = true
    fonts[#fonts + 1] = font_entry(id)
  end
end

-- the glyph's index in its font file (OpenType), else its slot
local function glyph_code(f, c)
  local ch = f and f.characters and f.characters[c]
  if ch and ch.index then return ch.index end
  return c
end

local out_glyph

-- a virtual font's character: its commands (dvi-like), from (x, y)
local function virtual_char(fid, f, c, x, y, glyphs, rules, fonts, depth)
  local ch = f.characters and f.characters[c]
  local commands = ch and ch.commands
  if not commands or depth > 8 then return end
  local size = f.size or 655360
  local scale = size / 1048576 -- VF units: 2^20 per design size (font.read_vf scales them already)
  local current = (f.fonts and f.fonts[1] and f.fonts[1].id) or fid
  local stack = {}
  for _, cmd in ipairs(commands) do
    local op = cmd[1]
    if op == "font" then
      local entry = f.fonts and f.fonts[cmd[2]]
      current = entry and entry.id or current
    elseif op == "char" or op == "slot" then
      local target, code = current, cmd[2]
      if op == "slot" then
        local entry = f.fonts and f.fonts[cmd[2]]
        target, code = entry and entry.id or current, cmd[3]
      end
      local tf = font.getfont(target)
      local width = out_glyph(target, tf, code, x, y, glyphs, rules, fonts, depth + 1)
      x = x + width
    elseif op == "right" then x = x + cmd[2]
    elseif op == "down" then y = y + cmd[2]
    elseif op == "rule" then
      local h, w = cmd[2], cmd[3]
      if w > 0 and h > 0 then rules[#rules + 1] = string.format("[%d,%d,%d,%d]", x, y - h, w, h) end
      x = x + w
    elseif op == "push" then stack[#stack + 1] = { x, y }
    elseif op == "pop" then
      local top = table.remove(stack)
      if top then x, y = top[1], top[2] end
    end
  end
end

-- a glyph at (x, y); its width
out_glyph = function(fid, f, c, x, y, glyphs, rules, fonts, depth)
  f = f or font.getfont(fid)
  local ch = f and f.characters and f.characters[c]
  if f and f.type == "virtual" and ch and ch.commands then
    virtual_char(fid, f, c, x, y, glyphs, rules, fonts, depth or 0)
    return ch.width or 0
  end
  note_font(fid, fonts)
  glyphs[#glyphs + 1] = string.format("[%d,%d,%d,%d]", fid, x, y, glyph_code(f, c))
  return ch and ch.width or 0
end

local walk_vlist

-- an hlist's contents, its baseline at y (y down), from x
local function walk_hlist(box, x, y, glyphs, rules, fonts, glues)
  local sign, order, set = box.glue_sign, box.glue_order, box.glue_set
  for n in node.traverse(box.head or box.list) do
    local id = n.id
    if id == GLYPH then
      local f = font.getfont(n.font)
      out_glyph(n.font, f, n.char, x + (n.xoffset or 0), y - (n.yoffset or 0), glyphs, rules, fonts, 0)
      x = x + n.width
    elseif id == HLIST then
      walk_hlist(n, x, y + (n.shift or 0), glyphs, rules, fonts, glues)
      x = x + n.width
    elseif id == VLIST then
      walk_vlist(n, x, y + (n.shift or 0) - n.height, glyphs, rules, fonts)
      x = x + n.width
    elseif id == RULE then
      local w = n.width
      local h = n.height == RUNNING and box.height or n.height
      local d = n.depth == RUNNING and box.depth or n.depth
      if w > 0 and h + d > 0 then rules[#rules + 1] = string.format("[%d,%d,%d,%d]", x, y - h, w, h + d) end
      x = x + w
    elseif id == GLUE then
      local w = n.width
      if sign == 1 and n.stretch_order == order then w = w + math.floor(n.stretch * set + 0.5)
      elseif sign == 2 and n.shrink_order == order then w = w - math.floor(n.shrink * set + 0.5) end
      -- (where it ends: SyncTeX records a glue there)
      if glues then glues[#glues + 1] = string.format("[%d,%d]", x + w, y) end
      if n.leader then
        -- a rule leader fills its glue; box leaders are drawn once (enough for a preview)
        local l = n.leader
        if l.id == RULE then
          local h = l.height == RUNNING and box.height or l.height
          local d = l.depth == RUNNING and box.depth or l.depth
          if w > 0 then rules[#rules + 1] = string.format("[%d,%d,%d,%d]", x, y - h, w, h + d) end
        end
      end
      x = x + w
    elseif id == KERN then
      x = x + n.kern
    elseif id == MATH then
      x = x + (n.surround or 0) + (n.width or 0)
    elseif id == DISC then
      -- a discretionary inside a line: its replacement stands there (after line breaking only the taken
      -- break's pre/post texts are left as plain nodes)
      local r = n.replace
      if r then
        local tmp = node.hpack(node.copy_list(r))
        walk_hlist(tmp, x, y, glyphs, rules, fonts, glues)
        x = x + tmp.width
        node.flush_node(tmp)
      end
    end
  end
end

-- a vlist's contents from its top at y; the lines (baseline, height, depth, width, shift) of its top level
walk_vlist = function(box, x, y, glyphs, rules, fonts, lines, glues)
  for n in node.traverse(box.head or box.list) do
    local id = n.id
    if id == HLIST then
      y = y + n.height
      if lines then lines[#lines + 1] = string.format("[%d,%d,%d,%d,%d]", y, n.height, n.depth, n.width, n.shift or 0) end
      walk_hlist(n, x + (n.shift or 0), y, glyphs, rules, fonts, glues)
      y = y + n.depth
    elseif id == VLIST then
      walk_vlist(n, x + (n.shift or 0), y, glyphs, rules, fonts)
      y = y + n.height + n.depth
    elseif id == GLUE then
      y = y + n.width
    elseif id == KERN then
      y = y + n.kern
    elseif id == RULE then
      local w = n.width == RUNNING and box.width or n.width
      if w > 0 then rules[#rules + 1] = string.format("[%d,%d,%d,%d]", x, y, w, n.height + n.depth) end
      y = y + n.height + n.depth
    end
  end
  return y
end

-- inserts (footnotes, floats), marks and adjusts the paragraph left in its box: it cannot be shown alone
local function has_migrated(box)
  for n in node.traverse(box.head or box.list) do
    if n.id == INS or n.id == ADJUST then return true end
    if (n.id == HLIST or n.id == VLIST) and has_migrated(n) then return true end
  end
  return false
end

-- TeX's errors while a paragraph is typeset (an undefined command half typed): counted, the first one said
local errors, first_error = 0, nil
luatexbase.add_to_callback("show_error_message", function()
  errors = errors + 1
  if not first_error then first_error = status.lasterrorstring or "error" end
  texio.write_nl("! " .. (status.lasterrorstring or ""))
end, "vtxlive.errors")

local function typeset(id, hsize, indent, leftskip, rightskip, text)
  local t0 = os.clock()
  local level = tex.currentgrouplevel
  errors, first_error = 0, nil
  tex.runtoks(function()
    tex.sprint(string.format("\\setbox0\\vbox{\\hsize=%dsp\\leftskip=%dsp\\rightskip=%dsp\\relax", hsize, leftskip, rightskip))
    if indent < -1 then
      -- the text says (the body's declarations come before its \indent or \noindent)
    elseif indent < 0 then
      tex.sprint("\\indent ")
    elseif indent == 0 then
      tex.sprint("\\noindent ")
    else
      tex.sprint(string.format("\\noindent\\hskip%dsp\\relax ", indent))
    end
    -- the paragraph's lines as TeX reads lines (each ends with \endlinechar: a comment ends there, an empty
    -- line is \par)
    local lines = {}
    for line in (text .. "\n"):gmatch("([^\n]*)\n") do lines[#lines + 1] = (line:gsub("\r$", "")) end
    tex.print(table.unpack(lines))
    tex.sprint("\\par}")
  end)
  if tex.currentgrouplevel ~= level then
    -- the paragraph left a group open (or closed one of the server's): this run's state is not the
    -- document's any more
    error("vtxlive: unbalanced groups")
  end
  local box = tex.box[0]
  local glyphs, rules, fonts, lines, glues = {}, {}, {}, {}, {}
  local total = walk_vlist(box, 0, 0, glyphs, rules, fonts, lines, glues)
  return string.format('{"id":"%s","ok":true,"ms":%.2f,"height":%d,"migrated":%s,"errors":%d,"error":"%s","lines":[%s],"glyphs":[%s],"rules":[%s],"glues":[%s],"fonts":[%s]}',
    esc(id), (os.clock() - t0) * 1000, total, tostring(has_migrated(box)), errors, esc(first_error or ""), table.concat(lines, ","),
    table.concat(glyphs, ","), table.concat(rules, ","), table.concat(glues, ","), table.concat(fonts, ","))
end

local function answer(json)
  io.write("\n@@VTXLIVE@@ ", json, "\n")
  io.flush()
end

answer('{"ready":true}')
while true do
  local head = io.read("l")
  if not head or head == "Q" then break end
  local id, hsize, indent, leftskip, rightskip, bytes = head:match("^P (%S+) (%-?%d+) (%-?%d+) (%-?%d+) (%-?%d+) (%d+)$")
  if id then
    local text = io.read(tonumber(bytes)) or ""
    local ok, result = pcall(typeset, id, tonumber(hsize), tonumber(indent), tonumber(leftskip), tonumber(rightskip), text)
    if not ok then result = string.format('{"id":"%s","ok":false,"error":"%s"}', esc(id), esc(tostring(result))) end
    answer(result)
  else
    local file = head:match("^R (.+)$")
    if file then
      -- the labels and citations of the document's last run (its .aux: \newlabel and \bibcite lines only; the
      -- rest of an .aux is for the preamble), into a box that is dropped
      local ok = pcall(function()
        local lines = {}
        local h = io.open(file, "r")
        if h then
          for line in h:lines() do
            if line:match("^\\newlabel{") or line:match("^\\bibcite{") then lines[#lines + 1] = line end
          end
          h:close()
        end
        if #lines > 0 then
          tex.runtoks(function()
            -- (\newlabel and \bibcite are the preamble's: their definitions here, as LaTeX's \@newl@bel makes them)
            tex.sprint("\\begingroup\\makeatletter\\def\\vtx@label#1#2#3{\\expandafter\\gdef\\csname #1@#2\\endcsname{#3}}")
            tex.sprint("\\def\\newlabel{\\vtx@label r}\\def\\bibcite{\\vtx@label b}\\setbox0\\vbox{")
            tex.print(table.unpack(lines))
            tex.sprint("}\\endgroup")
          end)
        end
      end)
      answer(string.format('{"reread":%s}', tostring(ok)))
    elseif head == "F" then
      -- the next answers say every font again (a new reader)
      fonts_sent = {}
      answer('{"fonts":"reset"}')
    end
  end
end
