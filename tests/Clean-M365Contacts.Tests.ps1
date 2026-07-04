#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

<#
    Unit tests for the pure, offline-testable functions in Clean-M365Contacts.ps1.

    These cover the logic that actually decides what gets deleted - key
    generation, completeness scoring, duplicate grouping, email extraction and
    JWT decoding - none of which touch the network. Anything that talks to Graph
    or the Windows credential store is intentionally not exercised here; that is
    validated by hand against a real mailbox.

    The script is dot-sourced with CLEANM365_NOAUTORUN set so only the functions
    load - the interactive menu does not run.
#>

BeforeAll {
    $env:CLEANM365_NOAUTORUN = '1'
    . (Join-Path (Join-Path $PSScriptRoot '..') 'Clean-M365Contacts.ps1')

    function New-TestContact {
        param(
            [string]$Id = [guid]::NewGuid().ToString(),
            [string]$DisplayName,
            [string]$GivenName,
            [string]$Surname,
            [string[]]$Emails = @(),
            [string]$Mobile,
            [string[]]$BusinessPhones = @(),
            [string]$Company,
            [string]$JobTitle
        )
        [pscustomobject]@{
            id             = $Id
            displayName    = $DisplayName
            givenName      = $GivenName
            surname        = $Surname
            emailAddresses = @($Emails | ForEach-Object { @{ address = $_ } })
            mobilePhone    = $Mobile
            businessPhones = $BusinessPhones
            companyName    = $Company
            jobTitle       = $JobTitle
        }
    }
}

Describe 'Get-FirstEmail' {
    It 'returns the first address, lower-cased' {
        $c = New-TestContact -Emails @('Foo.Bar@Example.COM')
        Get-FirstEmail $c | Should -BeExactly 'foo.bar@example.com'
    }

    It 'returns an empty string when there is no email' {
        $c = New-TestContact -DisplayName 'No Email'
        Get-FirstEmail $c | Should -BeExactly ''
    }
}

Describe 'Get-ContactKey' {
    BeforeAll {
        $script:c = New-TestContact -DisplayName 'Jane Doe' -Emails @('jane@example.com')
    }

    It 'NameEmail combines name and email' {
        Get-ContactKey $script:c 'NameEmail' | Should -BeExactly 'jane doe|jane@example.com'
    }

    It 'Name uses the display name only' {
        Get-ContactKey $script:c 'Name' | Should -BeExactly 'jane doe'
    }

    It 'Email uses the first email only' {
        Get-ContactKey $script:c 'Email' | Should -BeExactly 'jane@example.com'
    }

    It 'is whitespace and case insensitive on the name' {
        $c2 = New-TestContact -DisplayName '  JANE DOE ' -Emails @('jane@example.com')
        Get-ContactKey $c2 'NameEmail' | Should -BeExactly 'jane doe|jane@example.com'
    }
}

Describe 'Get-Score' {
    It 'scores a richer contact higher than a sparse one' {
        $rich   = New-TestContact -DisplayName 'Jane Doe' -GivenName 'Jane' -Surname 'Doe' `
                    -Emails @('jane@example.com') -Mobile '+441234' -Company 'Acme' -JobTitle 'CEO'
        $sparse = New-TestContact -DisplayName 'Jane Doe'
        (Get-Score $rich) | Should -BeGreaterThan (Get-Score $sparse)
    }

    It 'counts multiple emails' {
        $one = New-TestContact -DisplayName 'A' -Emails @('a@x.com')
        $two = New-TestContact -DisplayName 'A' -Emails @('a@x.com', 'a2@x.com')
        (Get-Score $two) | Should -BeGreaterThan (Get-Score $one)
    }
}

Describe 'Find-Duplicates' {
    It 'detects an exact name+email duplicate pair' {
        $contacts = @(
            New-TestContact -DisplayName 'Jane Doe' -Emails @('jane@example.com')
            New-TestContact -DisplayName 'Jane Doe' -Emails @('jane@example.com')
        )
        $dupes = Find-Duplicates $contacts 'NameEmail'
        $dupes.Count | Should -Be 1
        $dupes[0].Delete.Count | Should -Be 1
    }

    It 'keeps the most complete record and marks the stub for deletion' {
        $rich = New-TestContact -DisplayName 'Jane Doe' -Emails @('jane@example.com') `
                    -Mobile '+441234' -Company 'Acme'
        $stub = New-TestContact -DisplayName 'Jane Doe' -Emails @('jane@example.com')
        $dupes = Find-Duplicates @($stub, $rich) 'NameEmail'
        $dupes[0].Keep.id       | Should -BeExactly $rich.id
        $dupes[0].Delete[0].id  | Should -BeExactly $stub.id
    }

    It 'leaves unique contacts alone' {
        $contacts = @(
            New-TestContact -DisplayName 'Jane Doe'  -Emails @('jane@example.com')
            New-TestContact -DisplayName 'John Smith' -Emails @('john@example.com')
        )
        Find-Duplicates $contacts 'NameEmail' | Should -BeNullOrEmpty
    }

    It 'never flags contacts whose match key is empty (blank name, no email)' {
        $contacts = @(
            New-TestContact -Emails @()
            New-TestContact -Emails @()
        )
        Find-Duplicates $contacts 'NameEmail' | Should -BeNullOrEmpty
    }

    It 'finds more duplicates in the looser Name-only mode' {
        # Same person, two different emails: not caught by NameEmail, caught by Name
        $contacts = @(
            New-TestContact -DisplayName 'Jane Doe' -Emails @('jane@work.com')
            New-TestContact -DisplayName 'Jane Doe' -Emails @('jane@home.com')
        )
        (Find-Duplicates $contacts 'NameEmail') | Should -BeNullOrEmpty
        (Find-Duplicates $contacts 'Name').Count | Should -Be 1
    }
}

Describe 'Get-JwtClaims' {
    It 'decodes the payload of a JWT without validating the signature' {
        # header.payload.signature - payload is {"name":"Jane Doe","preferred_username":"jane@example.com"}
        $payload = '{"name":"Jane Doe","preferred_username":"jane@example.com"}'
        $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
        $b64url = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $jwt = "header.$b64url.signature"
        $claims = Get-JwtClaims $jwt
        $claims.name               | Should -BeExactly 'Jane Doe'
        $claims.preferred_username | Should -BeExactly 'jane@example.com'
    }

    It 'returns null on malformed input' {
        Get-JwtClaims 'not-a-jwt' | Should -BeNullOrEmpty
    }
}
