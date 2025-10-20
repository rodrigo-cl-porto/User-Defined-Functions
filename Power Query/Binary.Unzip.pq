let 
    func = (ZIPFile as binary) as table => let
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
        in Table.FromRecords(Entries),

    documentation = type function (
        ZIPFile as (type binary meta [
            Documentation.FieldCaption = "ZIP File",
            Documentation.FieldDescription = "The binary content of the ZIP file to unzip"
        ])) as (type table meta [
            Documentation.Name = "Binary.Unzip",
            Documentation.Description = "Unzips a binary ZIP file into a table with columns 'FileName' and 'Content'. Each row represents a file within the ZIP archive.",
            Documentation.LongDescription = "This function takes a binary representation of a ZIP file and extracts its contents into a table format. The resulting table contains two columns: 'FileName', which holds the name of each file in the ZIP archive, and 'Content', which contains the binary content of each respective file. The function supports standard ZIP compression and handles multiple files within the archive.",
            Documentation.Category = "Binary",
            Documentation.Version = "1.0.1",
            Documentation.Author = "Ignacio Barrau",
            Documentation.Examples = {[
                Description = "Unzips a ZIP file and retrieves the content of the first file.",
                Code = "let Files = Binary.Unzip(File.Contents(""C:\Temp\archive.zip"")) in Files{0}[Content]",
                Result = "binary content of the first file in the ZIP"
            ]}
        ]),
    documentedFunction = Value.ReplaceType(func, documentation)
in  
    documentedFunction
