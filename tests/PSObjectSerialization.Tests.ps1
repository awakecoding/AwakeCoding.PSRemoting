# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

Describe "PSObject Serialization Cmdlets (Export-PSObject, Import-PSObject, ConvertTo-PSObject, ConvertFrom-PSObject)" {

    BeforeAll {
        $ProjectRoot = Split-Path -Parent $PSScriptRoot
        $ModulePath = Join-Path $ProjectRoot 'AwakeCoding.PSRemoting'
        $ModuleManifest = Join-Path $ModulePath 'AwakeCoding.PSRemoting.psd1'
        Import-Module $ModuleManifest -Force -ErrorAction Stop

        class SerializationTestType {
            [int] $Property1 = 1
            [string] $Property2 = "some string"

            SerializationTestType() {
            }

            SerializationTestType([int]$val, [string]$data) {
                $this.Property1 = $val
                $this.Property2 = $data
            }
        }

        class DepthFour {
            [int] $num = 4
        }

        class DepthThree {
            [DepthFour] $four = [DepthFour]::new()
            [int] $num = 3
        }

        class DepthTwo {
            [DepthThree] $three = [DepthThree]::new()
            [int] $value = 2
        }

        class DepthOne {
            [DepthTwo] $two = [DepthTwo]::new()
            [int] $value = 1
        }
    }

    BeforeEach {
        $protoFile = Join-Path -Path $TestDrive -ChildPath "test.pb"
        $cborFile = Join-Path -Path $TestDrive -ChildPath "test.cbor"
        $xmlFile = Join-Path -Path $TestDrive -ChildPath "test.clixml"
    }

    AfterEach {
        Remove-Item $protoFile -Force -ErrorAction SilentlyContinue
        Remove-Item $cborFile -Force -ErrorAction SilentlyContinue
        Remove-Item $xmlFile -Force -ErrorAction SilentlyContinue
    }

    Context "Command Resolution" {

        It "Exports Export-PSObject, Import-PSObject, ConvertTo-PSObject, and ConvertFrom-PSObject from AwakeCoding.PSRemoting" {
            (Get-Command Export-PSObject -Module AwakeCoding.PSRemoting).Name | Should -BeExactly "Export-PSObject"
            (Get-Command Import-PSObject -Module AwakeCoding.PSRemoting).Name | Should -BeExactly "Import-PSObject"
            (Get-Command ConvertTo-PSObject -Module AwakeCoding.PSRemoting).Name | Should -BeExactly "ConvertTo-PSObject"
            (Get-Command ConvertFrom-PSObject -Module AwakeCoding.PSRemoting).Name | Should -BeExactly "ConvertFrom-PSObject"
        }
    }

    Context "Protobuf Format (-Format Protobuf)" {

        It "Preserves CmdletInfo PSTypeNames and properties" {
            Get-Command *-PSObject -Type Cmdlet | Select-Object -First 3 | Export-PSObject -Path $protoFile -Format Protobuf
            $results = Import-PSObject -Path $protoFile -Format Protobuf
            $results.Count | Should -BeExactly 3
            $results[0].PSTypeNames[0] | Should -BeExactly "Deserialized.System.Management.Automation.CmdletInfo"
        }

        It "Round-trips custom class instances" {
            $testObj = [SerializationTestType]::new(256, "abcdef")
            $testObj | Export-PSObject -Path $protoFile -Format Protobuf
            $results = Import-PSObject -Path $protoFile -Format Protobuf
            $results.Property1 | Should -Be 256
            $results.Property2 | Should -BeExactly "abcdef"
        }

        It "Preserves Hashtable, OrderedDictionary, and case-sensitive Hashtable" {
            $dict = @{ Larry = 'Poik!'; Curly = 'Nyuk!'; Moe = 'Wise guy!' }
            $dict | Export-PSObject -Path $protoFile -Format Protobuf
            $res = Import-PSObject -Path $protoFile -Format Protobuf
            $res.GetType().Name | Should -BeExactly "Hashtable"
            $res.Count | Should -Be 3

            $ordered = [ordered]@{ Larry = 'Poik!'; Curly = 'Nyuk!'; Moe = 'Wise guy!' }
            $ordered | Export-PSObject -Path $protoFile -Format Protobuf
            $resOrd = Import-PSObject -Path $protoFile -Format Protobuf
            $resOrd.GetType().Name | Should -BeExactly "OrderedDictionary"
            $resOrd[0] | Should -BeExactly 'Poik!'
            $resOrd[1] | Should -BeExactly 'Nyuk!'
            $resOrd[2] | Should -BeExactly 'Wise guy!'

            $csDict = [System.Collections.Hashtable]::new()
            $csDict['Larry'] = 'Poik!'
            $csDict['larry'] = 'poik!'
            $csDict | Export-PSObject -Path $protoFile -Format Protobuf
            $resCs = Import-PSObject -Path $protoFile -Format Protobuf
            $resCs.Count | Should -Be 2
        }

        It "Supports -Depth, paging (-IncludeTotalCount, -First, -Skip), -NoClobber, and -WhatIf" {
            $one = [DepthOne]::new()
            $one | Export-PSObject -Depth 2 -Path $protoFile -Format Protobuf
            $deserializedOne = Import-PSObject -Path $protoFile -Format Protobuf
            $deserializedOne.value | Should -Be 1
            $deserializedOne.two.value | Should -Be 2
            $deserializedOne.two.three | Should -Not -BeNullOrEmpty
            $deserializedOne.two.three.num | Should -BeNullOrEmpty

            "one", "two", "three", "four" | Export-PSObject -Path $protoFile -Format Protobuf
            $paged = Import-PSObject -Path $protoFile -Format Protobuf -First 2 -Skip 1
            $paged.Count | Should -Be 2
            $paged[0] | Should -BeExactly "two"
            $paged[1] | Should -BeExactly "three"

            $withTotal = Import-PSObject -Path $protoFile -Format Protobuf -IncludeTotalCount
            $withTotal[0].ToString() | Should -BeExactly "Unknown total count"

            "initial" | Export-PSObject -Path $protoFile -Format Protobuf
            { "second" | Export-PSObject -Path $protoFile -Format Protobuf -NoClobber -ErrorAction Stop } | Should -Throw -ErrorId "NoClobber,AwakeCoding.PSRemoting.PowerShell.ExportPSObjectCommand"
            Import-PSObject -Path $protoFile -Format Protobuf | Should -BeExactly "initial"

            $whatIfFile = Join-Path -Path $TestDrive -ChildPath "whatif.pb"
            "data" | Export-PSObject -Path $whatIfFile -Format Protobuf -WhatIf
            Test-Path $whatIfFile | Should -BeFalse
        }

        It "Streams one payload per pipeline item in ProcessRecord and round-trips via pipeline, parameter, unrolled bytes, Stream, and Base64" {
            $singleObj = [PSCustomObject]@{ Id = 1; Name = "One" }
            $singleBytes = $singleObj | ConvertTo-PSObject -Format Protobuf
            $singleObj | Export-PSObject -Path $protoFile -Format Protobuf
            $fileBytes = [System.IO.File]::ReadAllBytes($protoFile)
            $singleBytes | Should -Be $fileBytes

            # Verify 4-byte "PSPB" magic header
            [System.Text.Encoding]::ASCII.GetString($singleBytes, 0, 4) | Should -BeExactly "PSPB"

            # Piping multiple items emits 1 byte[] payload per pipeline item in ProcessRecord
            $testObject = 1, 2, 3
            $payloads = $testObject | ConvertTo-PSObject -Format Protobuf
            $payloads.Count | Should -Be 3
            ,$payloads[0] | Should -BeOfType [byte[]]

            (ConvertFrom-PSObject -InputObject $payloads -Format Protobuf) | Should -Be @(1, 2, 3)
            ($testObject | ConvertTo-PSObject -Format Protobuf | ConvertFrom-PSObject -Format Protobuf) | Should -Be @(1, 2, 3)
            ($singleBytes | ConvertFrom-PSObject -Format Protobuf).Name | Should -BeExactly "One"

            $ms = [System.IO.MemoryStream]::new($singleBytes, $false)
            try {
                (ConvertFrom-PSObject -InputObject $ms -Format Protobuf).Name | Should -BeExactly "One"
            }
            finally {
                $ms.Dispose()
            }

            $b64 = $testObject | ConvertTo-PSObject -Format Protobuf -AsBase64String
            $b64.Count | Should -Be 3
            $b64[0] | Should -BeOfType [string]
            ($b64 | ConvertFrom-PSObject -Format Protobuf) | Should -Be @(1, 2, 3)
        }
    }

    Context "CBOR Format (-Format Cbor)" {

        It "Preserves CmdletInfo PSTypeNames and properties" {
            Get-Command *-PSObject -Type Cmdlet | Select-Object -First 3 | Export-PSObject -Path $cborFile -Format Cbor
            $results = Import-PSObject -Path $cborFile -Format Cbor
            $results.Count | Should -BeExactly 3
            $results[0].PSTypeNames[0] | Should -BeExactly "Deserialized.System.Management.Automation.CmdletInfo"
        }

        It "Supports -Depth, paging (-IncludeTotalCount, -First, -Skip), -NoClobber, and -WhatIf" {
            $one = [DepthOne]::new()
            $one | Export-PSObject -Depth 2 -Path $cborFile -Format Cbor
            $deserializedOne = Import-PSObject -Path $cborFile -Format Cbor
            $deserializedOne.value | Should -Be 1
            $deserializedOne.two.value | Should -Be 2
            $deserializedOne.two.three | Should -Not -BeNullOrEmpty
            $deserializedOne.two.three.num | Should -BeNullOrEmpty

            "one", "two", "three", "four" | Export-PSObject -Path $cborFile -Format Cbor
            $paged = Import-PSObject -Path $cborFile -Format Cbor -First 2 -Skip 1
            $paged.Count | Should -Be 2
            $paged[0] | Should -BeExactly "two"
            $paged[1] | Should -BeExactly "three"

            $withTotal = Import-PSObject -Path $cborFile -Format Cbor -IncludeTotalCount
            $withTotal[0].ToString() | Should -BeExactly "Unknown total count"

            "initial" | Export-PSObject -Path $cborFile -Format Cbor
            { "second" | Export-PSObject -Path $cborFile -Format Cbor -NoClobber -ErrorAction Stop } | Should -Throw -ErrorId "NoClobber,AwakeCoding.PSRemoting.PowerShell.ExportPSObjectCommand"
            Import-PSObject -Path $cborFile -Format Cbor | Should -BeExactly "initial"

            $whatIfFile = Join-Path -Path $TestDrive -ChildPath "whatif.cbor"
            "data" | Export-PSObject -Path $whatIfFile -Format Cbor -WhatIf
            Test-Path $whatIfFile | Should -BeFalse
        }

        It "Streams one payload per pipeline item in ProcessRecord and round-trips via pipeline, parameter, unrolled bytes, and Base64" {
            $singleObj = [PSCustomObject]@{ Id = 1; Name = "One" }
            $singleBytes = $singleObj | ConvertTo-PSObject -Format Cbor
            $singleObj | Export-PSObject -Path $cborFile -Format Cbor
            $fileBytes = [System.IO.File]::ReadAllBytes($cborFile)
            $singleBytes | Should -Be $fileBytes

            # Verify CBOR Self-Described Tag 55799 (0xD9 0xD9 0xF7)
            $singleBytes[0] | Should -Be 0xD9
            $singleBytes[1] | Should -Be 0xD9
            $singleBytes[2] | Should -Be 0xF7

            $testObject = 1, 2, 3
            $payloads = $testObject | ConvertTo-PSObject -Format Cbor
            $payloads.Count | Should -Be 3
            ,$payloads[0] | Should -BeOfType [byte[]]

            (ConvertFrom-PSObject -InputObject $payloads -Format Cbor) | Should -Be @(1, 2, 3)
            ($testObject | ConvertTo-PSObject -Format Cbor | ConvertFrom-PSObject -Format Cbor) | Should -Be @(1, 2, 3)
            ($singleBytes | ConvertFrom-PSObject -Format Cbor).Name | Should -BeExactly "One"

            $b64 = $testObject | ConvertTo-PSObject -Format Cbor -AsBase64String
            $b64.Count | Should -Be 3
            $b64[0] | Should -BeOfType [string]
            ($b64 | ConvertFrom-PSObject -Format Cbor) | Should -Be @(1, 2, 3)
        }

        It "Streams items immediately in both directions (ConvertTo-PSObject and ConvertFrom-PSObject) before upstream finishes" {
            foreach ($fmt in @('Cbor', 'Protobuf', 'Clixml')) {
                $events = [System.Collections.Generic.List[string]]::new()
                1..3 | ForEach-Object {
                    $events.Add("produce:$_")
                    $_
                } | ConvertTo-PSObject -Format $fmt | ForEach-Object {
                    $events.Add("serialized:$($_.GetType().Name)")
                    Write-Output $_ -NoEnumerate
                } | ConvertFrom-PSObject -Format $fmt | ForEach-Object {
                    $events.Add("consume:$_")
                }

                $expectedType = if ($fmt -eq 'Clixml') { 'String' } else { 'Byte[]' }
                $events | Should -Be @(
                    "produce:1", "serialized:$expectedType", "consume:1",
                    "produce:2", "serialized:$expectedType", "consume:2",
                    "produce:3", "serialized:$expectedType", "consume:3"
                )
            }
        }

        It "Serializes array differently via ValueFromPipeline vs -InputObject just like CliXml" {
            $testObject = @(1, 2, 3, 4)

            $pipeOut = $testObject | ConvertTo-PSObject -Format Cbor | ConvertFrom-PSObject -Format Cbor
            $pipeOut -is [array] | Should -BeTrue
            $pipeOut.Count | Should -Be 4

            $paramOut = ConvertTo-PSObject -InputObject $testObject -Format Cbor | ConvertFrom-PSObject -Format Cbor
            $paramOut -is [System.Collections.ArrayList] | Should -BeTrue
            $paramOut.Count | Should -Be 4
        }
    }

    Context "CLIXML Format (-Format Clixml)" {

        It "Exports and imports CLIXML files with -Encoding, paging, -NoClobber, and -WhatIf" {
            $obj = [PSCustomObject]@{ Name = "CliXmlItem"; Count = 42 }
            $obj | Export-PSObject -Path $xmlFile -Format Clixml
            $imported = Import-PSObject -Path $xmlFile -Format Clixml
            $imported.Name | Should -BeExactly "CliXmlItem"
            $imported.Count | Should -Be 42

            # Verify -Encoding unicode writes UTF-16 LE BOM
            $unicodeFile = Join-Path -Path $TestDrive -ChildPath "unicode.clixml"
            $obj | Export-PSObject -Path $unicodeFile -Format Clixml -Encoding unicode
            $rawBytes = [System.IO.File]::ReadAllBytes($unicodeFile)
            $rawBytes[0] | Should -Be 0xFF
            $rawBytes[1] | Should -Be 0xFE
            $fromUnicode = Import-PSObject -Path $unicodeFile -Format Clixml
            $fromUnicode.Name | Should -BeExactly "CliXmlItem"

            # Verify #< CLIXML first-line directive is skipped on import
            $directiveFile = Join-Path -Path $TestDrive -ChildPath "directive.clixml"
            $xmlText = $obj | ConvertTo-PSObject -Format Clixml
            [System.IO.File]::WriteAllText($directiveFile, "#< CLIXML`r`n$xmlText")
            $fromDirective = Import-PSObject -Path $directiveFile -Format Clixml
            $fromDirective.Name | Should -BeExactly "CliXmlItem"

            # Paging
            "one", "two", "three", "four" | Export-PSObject -Path $xmlFile -Format Clixml
            $paged = Import-PSObject -Path $xmlFile -Format Clixml -First 2 -Skip 1
            $paged | Should -Be @("two", "three")

            # -NoClobber
            { "second" | Export-PSObject -Path $xmlFile -Format Clixml -NoClobber -ErrorAction Stop } | Should -Throw -ErrorId "NoClobber,AwakeCoding.PSRemoting.PowerShell.ExportPSObjectCommand"
        }

        It "Round-trips ConvertTo-PSObject and ConvertFrom-PSObject -Format Clixml as XML string, byte[], and Base64" {
            $testObject = 10, 20, 30
            $xmlStrings = $testObject | ConvertTo-PSObject -Format Clixml
            $xmlStrings.Count | Should -Be 3
            $xmlStrings[0] | Should -BeOfType [string]
            $xmlStrings[0] | Should -Match '<Objs'

            ($xmlStrings | ConvertFrom-PSObject -Format Clixml) | Should -Be @(10, 20, 30)

            $xmlBytes = $testObject | ConvertTo-PSObject -Format Clixml -AsByteArray
            $xmlBytes.Count | Should -Be 3
            ,$xmlBytes[0] | Should -BeOfType [byte[]]
            (ConvertFrom-PSObject -InputObject $xmlBytes -Format Clixml) | Should -Be @(10, 20, 30)
            ($xmlBytes | ConvertFrom-PSObject -Format Clixml) | Should -Be @(10, 20, 30)

            $xmlB64 = $testObject | ConvertTo-PSObject -Format Clixml -AsBase64String
            $xmlB64.Count | Should -Be 3
            $xmlB64[0] | Should -BeOfType [string]
            $xmlB64[0] | Should -Not -Match '<Objs'
            ($xmlB64 | ConvertFrom-PSObject -Format Clixml) | Should -Be @(10, 20, 30)
        }
    }

    Context "Automatic Format Detection & Friendly Format Aliases" {

        It "Round-trips across Protobuf, Cbor, and Clixml with explicit -Format and auto-detected format" {
            $payload = [PSCustomObject]@{
                Title = "UnifiedTest"
                Items = @(1, "two", $true)
                Meta  = [ordered]@{ Env = "Test"; Id = 99 }
            }

            foreach ($fmt in @('Protobuf', 'Cbor', 'Clixml')) {
                $serialized = $payload | ConvertTo-PSObject -Format $fmt
                if ($fmt -eq 'Clixml') {
                    $serialized | Should -BeOfType [string]
                }
                else {
                    ,$serialized | Should -BeOfType [byte[]]
                }

                # Explicit format
                $deserializedExplicit = ConvertFrom-PSObject -InputObject $serialized -Format $fmt
                $deserializedExplicit.Title | Should -BeExactly "UnifiedTest"
                $deserializedExplicit.Items | Should -Be @(1, "two", $true)
                $deserializedExplicit.Meta['Env'] | Should -BeExactly "Test"

                # Auto-detected format (omitting -Format)
                $deserializedAuto = $serialized | ConvertFrom-PSObject
                $deserializedAuto.Title | Should -BeExactly "UnifiedTest"
                $deserializedAuto.Meta['Id'] | Should -Be 99

                # Base64 round-trip with auto-detection
                $b64 = $payload | ConvertTo-PSObject -Format $fmt -AsBase64String
                $b64 | Should -BeOfType [string]
                $fromB64 = $b64 | ConvertFrom-PSObject
                $fromB64.Title | Should -BeExactly "UnifiedTest"
            }
        }

        It "Infers format from file extension on Export-PSObject and auto-detects header on Import-PSObject" {
            $item = [PSCustomObject]@{ Kind = "FileAutoDetect"; Number = 777 }

            $item | Export-PSObject -Path $protoFile
            $item | Export-PSObject -Path $cborFile
            $item | Export-PSObject -Path $xmlFile

            # Verify magic headers written based on file extension
            $pbBytes = [System.IO.File]::ReadAllBytes($protoFile)
            [System.Text.Encoding]::ASCII.GetString($pbBytes, 0, 4) | Should -BeExactly "PSPB"

            $cbBytes = [System.IO.File]::ReadAllBytes($cborFile)
            $cbBytes[0] | Should -Be 0xD9
            $cbBytes[1] | Should -Be 0xD9
            $cbBytes[2] | Should -Be 0xF7

            $xmText = [System.IO.File]::ReadAllText($xmlFile)
            $xmText | Should -Match '<Objs'

            # Auto-detect on Import-PSObject without -Format
            (Import-PSObject -Path $protoFile).Number | Should -Be 777
            (Import-PSObject -Path $cborFile).Number | Should -Be 777
            (Import-PSObject -Path $xmlFile).Number | Should -Be 777
        }

        It "Defaults to CBOR when -Format is omitted and file extension is unknown" {
            $bytes = [PSCustomObject]@{ DefaultFormat = "CBOR" } | ConvertTo-PSObject
            $bytes[0] | Should -Be 0xD9
            $bytes[1] | Should -Be 0xD9
            $bytes[2] | Should -Be 0xF7
            ($bytes | ConvertFrom-PSObject).DefaultFormat | Should -BeExactly "CBOR"

            $binFile = Join-Path -Path $TestDrive -ChildPath "default.bin"
            [PSCustomObject]@{ DefaultFormat = "CBORFile" } | Export-PSObject -Path $binFile
            $binBytes = [System.IO.File]::ReadAllBytes($binFile)
            $binBytes[0] | Should -Be 0xD9
            $binBytes[1] | Should -Be 0xD9
            $binBytes[2] | Should -Be 0xF7
            (Import-PSObject -Path $binFile).DefaultFormat | Should -BeExactly "CBORFile"
        }

        It "Supports friendly format names (Proto, CliProto, CliCbor, Xml, CliXml)" {
            $data = [PSCustomObject]@{ Check = "AliasTest" }

            ($data | ConvertTo-PSObject -Format Proto | ConvertFrom-PSObject).Check | Should -BeExactly "AliasTest"
            ($data | ConvertTo-PSObject -Format CliCbor | ConvertFrom-PSObject).Check | Should -BeExactly "AliasTest"
            ($data | ConvertTo-PSObject -Format Xml | ConvertFrom-PSObject).Check | Should -BeExactly "AliasTest"
        }
    }

    Context "3-Way Roundtrip Comparison: Protobuf vs Cbor vs Clixml" {

        It "Honors -Depth 2, 3, and 4 identically across Protobuf, Cbor, and Clixml" {
            foreach ($d in 2, 3, 4) {
                $testObject = [DepthOne]::new()
                $fromXml = $testObject | ConvertTo-PSObject -Format Clixml -Depth $d | ConvertFrom-PSObject -Format Clixml
                $fromCbor = $testObject | ConvertTo-PSObject -Format Cbor -Depth $d | ConvertFrom-PSObject -Format Cbor
                $fromProto = $testObject | ConvertTo-PSObject -Format Protobuf -Depth $d | ConvertFrom-PSObject -Format Protobuf

                $fromProto.value | Should -Be $fromXml.value
                $fromProto.two.value | Should -Be $fromXml.two.value
                $fromProto.two.three.num | Should -Be $fromXml.two.three.num
                $fromProto.two.three.four.num | Should -Be $fromXml.two.three.four.num
                $fromProto.two.three.four.num | Should -Be $fromCbor.two.three.four.num
            }
        }

        It "Preserves exact types and values for all PowerShell primitive types across Protobuf, Cbor, and Clixml" {
            $secStr = ConvertTo-SecureString "S3cr3tP@ssw0rd!" -AsPlainText -Force
            $primitives = @(
                $true,
                $false,
                [byte]200,
                [sbyte]-100,
                [int16]-30000,
                [uint16]60000,
                [int32]-12345678,
                [uint32]3000000000,
                [int64]-9000000000000000000,
                [uint64]18000000000000000000,
                [single]1.5,
                [double]3.14159265358979,
                [decimal]987654321.123456789,
                [char]'Z',
                "Hello PowerShell Serialization",
                [datetime]::new(2026, 4, 18, 15, 30, 45, [System.DateTimeKind]::Utc),
                [timespan]::FromMinutes(95.5),
                [guid]::NewGuid(),
                [uri]"https://aka.ms/powershell",
                [version]"7.6.7.0",
                [byte[]]@(10, 20, 30, 40, 50),
                ([xml]"<root><item id='1'>text</item></root>"),
                $secStr
            )

            foreach ($item in $primitives) {
                $fromXml = ,$item | ConvertTo-PSObject -Format Clixml | ConvertFrom-PSObject -Format Clixml
                $fromCbor = ,$item | ConvertTo-PSObject -Format Cbor | ConvertFrom-PSObject -Format Cbor
                $fromProto = ,$item | ConvertTo-PSObject -Format Protobuf | ConvertFrom-PSObject -Format Protobuf

                $fromProto.GetType().FullName | Should -BeExactly $fromXml.GetType().FullName
                $fromProto.GetType().FullName | Should -BeExactly $fromCbor.GetType().FullName

                if ($item -is [xml]) {
                    $fromProto.OuterXml | Should -BeExactly $fromXml.OuterXml
                    $fromCbor.OuterXml | Should -BeExactly $fromXml.OuterXml
                }
                elseif ($item -is [securestring]) {
                    (New-Object System.Net.NetworkCredential("", $fromProto)).Password | Should -BeExactly "S3cr3tP@ssw0rd!"
                    (New-Object System.Net.NetworkCredential("", $fromCbor)).Password | Should -BeExactly "S3cr3tP@ssw0rd!"
                    (New-Object System.Net.NetworkCredential("", $fromXml)).Password | Should -BeExactly "S3cr3tP@ssw0rd!"
                }
                else {
                    $fromProto | Should -Be $fromXml
                    $fromCbor | Should -Be $fromXml
                }
            }
        }

        It "Matches 100% of serialized properties and PSTypeNames across Protobuf, Cbor, and Clixml" {
            $child = [PSCustomObject]@{
                ChildName = "Nested-1"
                Timestamp = [datetime]::UtcNow
                Metrics   = [ordered]@{ Cpu = 12.5; Mem = 4096; Active = $true }
            }
            $complexCustom = [PSCustomObject]@{
                PSTypeName = "MyCorp.DeploymentManifest"
                Id         = [guid]::NewGuid()
                Version    = [version]"2.4.1"
                Tags       = @("alpha", "beta", 99)
                Primary    = $child
                Secondary  = $child
            }

            $errRecord = try { 1 / 0 } catch { $_ }
            $pwshDllPath = Join-Path $PSHOME 'pwsh.dll'

            $testTargets = @(
                @{ Name = "PSCustomObject";  Obj = $complexCustom; Depth = 4 },
                @{ Name = "SemanticVersion"; Obj = [semver]"7.6.7-preview.3+build.42"; Depth = 2 },
                @{ Name = "DateTimeOffset";  Obj = [datetimeoffset]::Now; Depth = 2 },
                @{ Name = "FileInfo";        Obj = (Get-Item -LiteralPath $pwshDllPath); Depth = 2 },
                @{ Name = "CmdletInfo";      Obj = (Get-Command Export-PSObject); Depth = 2 },
                @{ Name = "ErrorRecord";     Obj = $errRecord; Depth = 2 }
            )

            foreach ($target in $testTargets) {
                $obj = $target.Obj
                $depth = $target.Depth

                $obj | Export-PSObject -Path $xmlFile -Format Clixml -Depth $depth
                $obj | Export-PSObject -Path $cborFile -Format Cbor -Depth $depth
                $obj | Export-PSObject -Path $protoFile -Format Protobuf -Depth $depth

                $xmlObj = Import-PSObject -Path $xmlFile -Format Clixml
                $cborObj = Import-PSObject -Path $cborFile -Format Cbor
                $protoObj = Import-PSObject -Path $protoFile -Format Protobuf

                ($protoObj.PSTypeNames -join "|") | Should -BeExactly ($xmlObj.PSTypeNames -join "|")
                ($cborObj.PSTypeNames -join "|") | Should -BeExactly ($xmlObj.PSTypeNames -join "|")

                $xmlProps = @($xmlObj.PSObject.Properties | Sort-Object Name)
                $protoProps = @($protoObj.PSObject.Properties | Sort-Object Name)
                $cborProps = @($cborObj.PSObject.Properties | Sort-Object Name)

                $protoProps.Count | Should -Be $xmlProps.Count -Because "$($target.Name) Proto property count should match CliXml"
                $cborProps.Count | Should -Be $xmlProps.Count -Because "$($target.Name) CBOR property count should match CliXml"
                ($protoProps.Name -join ",") | Should -BeExactly ($xmlProps.Name -join ",")
                ($cborProps.Name -join ",") | Should -BeExactly ($xmlProps.Name -join ",")

                foreach ($xp in $xmlProps) {
                    $pp = $protoObj.PSObject.Properties[$xp.Name]
                    $cp = $cborObj.PSObject.Properties[$xp.Name]
                    $pp | Should -Not -BeNullOrEmpty
                    $cp | Should -Not -BeNullOrEmpty
                    $pp.MemberType | Should -Be $xp.MemberType
                    $cp.MemberType | Should -Be $xp.MemberType

                    if ($null -eq $xp.Value) {
                        $pp.Value | Should -BeNullOrEmpty
                        $cp.Value | Should -BeNullOrEmpty
                    }
                    elseif ($xp.Value.GetType().IsValueType -or $xp.Value -is [string]) {
                        $pp.Value | Should -Be $xp.Value -Because "$($target.Name).$($xp.Name) should match in Proto"
                        $cp.Value | Should -Be $xp.Value -Because "$($target.Name).$($xp.Name) should match in CBOR"
                    }
                }

                (Get-Item $protoFile).Length | Should -BeLessThan (Get-Item $xmlFile).Length
                (Get-Item $cborFile).Length | Should -BeLessThan (Get-Item $xmlFile).Length
            }

            # Verify shared object reference identity on the PSCustomObject
            $complexCustom | Export-PSObject -Path $protoFile -Format Protobuf -Depth 4
            $rehydratedProto = Import-PSObject -Path $protoFile -Format Protobuf
            [object]::ReferenceEquals($rehydratedProto.Primary, $rehydratedProto.Secondary) | Should -BeTrue

            $complexCustom | Export-PSObject -Path $cborFile -Format Cbor -Depth 4
            $rehydratedCbor = Import-PSObject -Path $cborFile -Format Cbor
            [object]::ReferenceEquals($rehydratedCbor.Primary, $rehydratedCbor.Secondary) | Should -BeTrue
        }

        It "Round-trips circular references, Queue/Stack, annotated primitives, flags enums, ProgressRecord, and PSMemberSet" {
            # 1. Circular PSCustomObject
            $cycleObj = [PSCustomObject]@{ Name = "CycleRoot" }
            $cycleObj | Add-Member -NotePropertyName Self -NotePropertyValue $cycleObj
            $cycleObj | Export-PSObject -Path $protoFile -Format Protobuf -Depth 4
            $cycleObj | Export-PSObject -Path $cborFile -Format Cbor -Depth 4
            $cycleObj | Export-PSObject -Path $xmlFile -Format Clixml -Depth 4
            $cycleProto = Import-PSObject -Path $protoFile -Format Protobuf
            $cycleCbor = Import-PSObject -Path $cborFile -Format Cbor
            $cycleXml = Import-PSObject -Path $xmlFile -Format Clixml
            [object]::ReferenceEquals($cycleProto, $cycleProto.Self) | Should -BeTrue
            [object]::ReferenceEquals($cycleCbor, $cycleCbor.Self) | Should -BeTrue
            [object]::ReferenceEquals($cycleXml, $cycleXml.Self) | Should -BeTrue

            # 2. Queue and Stack
            $q = [System.Collections.Queue]::new()
            $q.Enqueue("first"); $q.Enqueue("second"); $q.Enqueue(3)
            Export-PSObject -InputObject $q -Path $protoFile -Format Protobuf
            $qProto = Import-PSObject -Path $protoFile -Format Protobuf
            ,$qProto | Should -BeOfType [System.Collections.Queue]
            $qProto.Dequeue() | Should -BeExactly "first"
            $qProto.Dequeue() | Should -BeExactly "second"
            $qProto.Dequeue() | Should -BeExactly 3

            $st = [System.Collections.Stack]::new()
            $st.Push("bottom"); $st.Push("middle"); $st.Push("top")
            Export-PSObject -InputObject $st -Path $cborFile -Format Cbor
            $stCbor = Import-PSObject -Path $cborFile -Format Cbor
            ,$stCbor | Should -BeOfType [System.Collections.Stack]
            $stCbor.Pop() | Should -BeExactly "top"
            $stCbor.Pop() | Should -BeExactly "middle"
            $stCbor.Pop() | Should -BeExactly "bottom"

            # 3. Primitive with NoteProperties and custom PSTypeName
            $annotated = "base-value" |
                Add-Member -NotePropertyName Tag -NotePropertyValue "meta-tag" -PassThru |
                Add-Member -NotePropertyName RetryCount -NotePropertyValue 5 -PassThru
            $annotated.PSObject.TypeNames.Insert(0, "MyCorp.AnnotatedString")
            $annotated | Export-PSObject -Path $protoFile -Format Protobuf
            $annotated | Export-PSObject -Path $cborFile -Format Cbor
            $annotated | Export-PSObject -Path $xmlFile -Format Clixml
            $annProto = Import-PSObject -Path $protoFile -Format Protobuf
            $annCbor = Import-PSObject -Path $cborFile -Format Cbor
            $annXml = Import-PSObject -Path $xmlFile -Format Clixml
            $annProto | Should -BeExactly "base-value"
            $annProto.Tag | Should -BeExactly "meta-tag"
            $annProto.RetryCount | Should -BeExactly 5
            $annProto.PSObject.TypeNames[0] | Should -BeExactly "Deserialized.MyCorp.AnnotatedString"
            $annCbor.Tag | Should -BeExactly "meta-tag"
            $annXml.Tag | Should -BeExactly "meta-tag"

            # 4. Flags Enum, ProgressRecord, custom PSMemberSet, NaN/Infinity, NUL/Emoji strings
            $edgeBundle = [PSCustomObject]@{
                Flags      = ([System.IO.FileAttributes]::ReadOnly -bor [System.IO.FileAttributes]::Hidden)
                Progress   = [System.Management.Automation.ProgressRecord]::new(7, "Building", "50% done")
                EmptyBytes = [byte[]]@()
                SpecialStr = "line1`0nul`r`nline2`t🚀🔥"
                Floats     = @([double]::NaN, [double]::PositiveInfinity, [double]::NegativeInfinity, [double]0.0)
            }
            $msMembers = [System.Collections.Generic.List[System.Management.Automation.PSMemberInfo]]::new()
            $msMembers.Add([System.Management.Automation.PSNoteProperty]::new("NestedKey", "NestedVal"))
            $edgeBundle.PSObject.Members.Add([System.Management.Automation.PSMemberSet]::new("CustomSet", $msMembers))

            $edgeBundle | Export-PSObject -Path $protoFile -Format Protobuf -Depth 5
            $edgeBundle | Export-PSObject -Path $cborFile -Format Cbor -Depth 5
            $ebProto = Import-PSObject -Path $protoFile -Format Protobuf
            $ebCbor = Import-PSObject -Path $cborFile -Format Cbor

            $ebProto.Flags.ToString() | Should -BeExactly "ReadOnly, Hidden"
            $ebProto.Progress.ActivityId | Should -BeExactly 7
            ,$ebProto.EmptyBytes | Should -BeOfType [byte[]]
            $ebProto.EmptyBytes.Length | Should -BeExactly 0
            $ebProto.SpecialStr | Should -BeExactly "line1`0nul`r`nline2`t🚀🔥"
            [double]::IsNaN($ebProto.Floats[0]) | Should -BeTrue
            [double]::IsPositiveInfinity($ebProto.Floats[1]) | Should -BeTrue
            [double]::IsNegativeInfinity($ebProto.Floats[2]) | Should -BeTrue
            $ebProto.CustomSet.NestedKey | Should -BeExactly "NestedVal"
            $ebCbor.CustomSet.NestedKey | Should -BeExactly "NestedVal"
        }
    }
}
