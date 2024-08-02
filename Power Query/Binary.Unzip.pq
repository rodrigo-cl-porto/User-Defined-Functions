(ZIPFile as binary) as table => 
let
    // Predefined byteformats that are used many times over
    Unsigned16BitLittleIEndian = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian),
    Unsigned32BitLittleIEndian = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
    Header = BinaryFormat.Record([
        MiscHeader = BinaryFormat.Binary(14),
        BinarySize = Unsigned32BitLittleIEndian,
        FileSize   = Unsigned32BitLittleIEndian,
        FileNameLen= Unsigned16BitLittleIEndian,
        ExtrasLen  = Unsigned16BitLittleIEndian    
    ]),
    HeaderChoice = BinaryFormat.Choice(
        Unsigned32BitLittleIEndian,
        each if _ <> 67324752 then  // not the IsValid number? then return a dummy formatter
            BinaryFormat.Record([IsValid = false, Filename=null, Content=null])
        else
            BinaryFormat.Choice(
                BinaryFormat.Binary(26), // Header payload - 14+4+4+2+2
                each BinaryFormat.Record([
                    IsValid  = true,
                    Filename = BinaryFormat.Text(Header(_)[FileNameLen]), 
                    Extras   = BinaryFormat.Text(Header(_)[ExtrasLen]), 
                    Content  = BinaryFormat.Transform(
                        BinaryFormat.Binary(Header(_)[BinarySize]),
                        (x) => try Binary.Buffer(Binary.Decompress(x, Compression.Deflate)) otherwise null
                    )
                ]),
                type binary                   // enable streaming
            )
    ),
    ZipFormat = BinaryFormat.List(HeaderChoice, each _[IsValid] = true),
    Entries = List.Transform(
        List.RemoveLastN(ZipFormat(ZIPFile), 1),
        (e) => [FileName = e[Filename], Content = e[Content]]
    )
in
    Table.FromRecords(Entries)
