@{
    # Run every built-in rule...
    IncludeDefaultRules = $true

    # ...except a small, deliberately-chosen set that does not fit an
    # interactive, single-file console tool. Each exclusion is justified below
    # so reviewers can see this is intentional, not a way to hide real problems.
    ExcludeRules = @(
        # This is an interactive menu. Write-Host is the correct way to paint a
        # console UI; its output is meant for the screen, not the pipeline.
        'PSAvoidUsingWriteHost',

        # The refresh token is wrapped with DPAPI (ConvertFrom-SecureString) and
        # stored per-user/per-machine. ConvertTo-SecureString -AsPlainText is the
        # documented way to seed that SecureString from an already-in-memory token.
        'PSAvoidUsingConvertToSecureStringWithPlainText',

        # This is a standalone script, not a module exporting cmdlets. Deletion is
        # already gated behind an explicit typed "DELETE" confirmation and a dry
        # run, so -WhatIf/-Confirm plumbing on internal helpers adds no safety.
        'PSUseShouldProcessForStateChangingFunctions',

        # Plural nouns (Get-AllContacts, Find-Duplicates, Get-JwtClaims) read more
        # naturally here and describe collections accurately.
        'PSUseSingularNouns'
    )
}
