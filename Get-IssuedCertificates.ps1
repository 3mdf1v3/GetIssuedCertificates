[CmdletBinding()]

Param (
 
    # Maximum number of days from now that a certificate will expire. (Default: 21900 = 60 years)
    [Int]
    $ExpireInDays = 21900,

    # Certificate Authority location string "computername\CAName" (Default gets location strings from Current Domain)
    [String[]]
    $CAlocation = (""),

    # Fields in the Certificate Authority Database to Export
    [String[]]
    $Properties = (
        'Issued Request ID',
        'Issued Common Name',
        'Certificate Expiration Date',
        'Certificate Effective Date',
        'Certificate Template',
        'Certificate Hash'
        # 'Request Disposition',
        # 'Issued Email Address',
        # 'Request Disposition Message',
        # 'Binary Certificate'
        # 'Requester Name'
        ),

    # Filter on Certificate Template OID (use Get-CertificateTemplateOID)
    [AllowNull()]
    [String]
    $CertificateTemplateOid,

    # Filter by Issued Common Name
    [AllowNull()]
    [String]
    $CommonName,

    # Filter by Issued Request ID
    [Int]
    $RequestId,
 
    [Int]
    $AsParam = 0

)

 

foreach ($Location in $CAlocation)
{

    $CaView = New-Object -ComObject CertificateAuthority.View
    $null = $CaView.OpenConnection($Location)
    $CaView.SetResultColumnCount($Properties.Count)

    #region SetOutput Colum
    foreach ($item in $Properties)
    {

        $index = $CaView.GetColumnIndex($false, $item)
        $CaView.SetResultColumn($index)

    }
    #endregion

    #region Filters
    $CVR_SEEK_EQ = 1
    $CVR_SEEK_LT = 2
    $CVR_SEEK_GT = 16

    #region filter expiration Date
    $index = $CaView.GetColumnIndex($false, 'Certificate Expiration Date')
    $now = Get-Date
    $expirationdate = $now.AddDays($ExpireInDays)

    if ($ExpireInDays -gt 0)
    {

        $CaView.SetRestriction($index,$CVR_SEEK_GT,0,$now)
        $CaView.SetRestriction($index,$CVR_SEEK_LT,0,$expirationdate)

    }
    else
    {

        $CaView.SetRestriction($index,$CVR_SEEK_LT,0,$now)
        $CaView.SetRestriction($index,$CVR_SEEK_GT,0,$expirationdate)

    }
    #endregion filter expiration date

    #region Filter Template
    if ($CertificateTemplateOid)
    {

        $index = $CaView.GetColumnIndex($false, 'Certificate Template')
        $CaView.SetRestriction($index,$CVR_SEEK_EQ,0,$CertificateTemplateOid)

    }
    #endregion

    #region Filter Issued Common Name
    if ($CommonName)
    {

        $index = $CaView.GetColumnIndex($false, 'Issued Common Name')
        $CaView.SetRestriction($index,$CVR_SEEK_EQ,0,$CommonName)

    }
    #endregion

    #region Filter Issued Request ID
    if ($RequestId)
    {
        $index = $CaView.GetColumnIndex($false, 'Issued Request ID')
        $CaView.SetRestriction($index,$CVR_SEEK_EQ,0,$RequestId)
    }
    #endregion

    # With Not Issued Common Name null
    $CaView.SetRestriction($CaView.GetColumnIndex($false, 'Issued Common Name'),$CVR_SEEK_LT,0,'null')
    #region Filter Only issued certificates

    # 20 - issued certificates
    $CaView.SetRestriction($CaView.GetColumnIndex($false, 'Request Disposition'),$CVR_SEEK_EQ,0,20)

    #endregion

    #endregion

    #region output each retuned row

    $CV_OUT_BASE64HEADER = 0
    $CV_OUT_BASE64 = 1

    $Now = get-Date
    $ExpirationDate = $now.AddDays(15)

    $RowObj = $CaView.OpenView()
    $certArr = @()

    while ($RowObj.Next() -ne -1)
    {

        $Cert = New-Object -TypeName PsObject
        $ColObj = $RowObj.EnumCertViewColumn()
        $null = $ColObj.Next()
        do
        {

            $displayName = $ColObj.GetDisplayName()
            if ($displayName -eq 'Binary Certificate')
            {

                $Cert | Add-Member -MemberType NoteProperty -Name $displayName -Value $($ColObj.GetValue($CV_OUT_BASE64HEADER)) -Force

            } elseif ($displayName -eq 'Issued Common Name')
            {

                if ( $AsParam -eq 1)
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name '{#CERT_ISSUED_COMMON_NAME}' -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force

                } else {

                    $Cert | Add-Member -MemberType NoteProperty -Name 'CERT_ISSUED_COMMON_NAME' -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force

                }            

            } elseif ($displayName -eq 'Certificate Expiration Date') 
            {

                $DateDiff = New-TimeSpan -Start ($Now) -End ($($ColObj.GetValue($CV_OUT_BASE64)))
                if ( $AsParam -eq 1)
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name '{#CERT_EXPIRATION_DATE}' -Value $($ColObj.GetValue($CV_OUT_BASE64).ToString("dd.MM.yyyy hh:mm:ss")) -Force
                    $Cert | Add-Member -MemberType NoteProperty -Name '{#CERT_DAYS_TO_EXPIRE}' -Value $datediff.Days -Force

                } else
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name 'CERT_EXPIRATION_DATE' -Value $($ColObj.GetValue($CV_OUT_BASE64).ToString("dd.MM.yyyy hh:mm:ss")) -Force
                    $Cert | Add-Member -MemberType NoteProperty -Name 'CERT_DAYS_TO_EXPIRE' -Value $datediff.Days -Force

                }                             

            } elseif ($displayName -eq 'Certificate Effective Date')
            {

                if ( $AsParam -eq 1)
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name '{#CERT_EFFECTIVE_DATE}' -Value $($ColObj.GetValue($CV_OUT_BASE64).ToString("dd.MM.yyyy hh:mm:ss")) -Force

                } else
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name 'CERT_EFFECTIVE_DATE' -Value $($ColObj.GetValue($CV_OUT_BASE64).ToString("dd.MM.yyyy hh:mm:ss")) -Force

                }

            } elseif ($displayName -eq 'Issued Request ID')
            {

                if ( $AsParam -eq 1)
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name '{#CERT_REQUEST_ID}' -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force

                } else
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name 'CERT_REQUEST_ID' -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force

                }

            } elseif ($displayName -eq 'Certificate Hash')
            {

                if ( $AsParam -eq 1)
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name '{#CERT_HASH}' -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force

                } else {

                    $Cert | Add-Member -MemberType NoteProperty -Name 'CERT_HASH' -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force

                }

            } elseif ($displayName -eq 'Certificate Template')  
            {

                if ( $AsParam -eq 1)
                {

                    $Cert | Add-Member -MemberType NoteProperty -Name '{#CERT_CERTIFICATE_TEMPLATE}' -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force

                } else {

                    $Cert | Add-Member -MemberType NoteProperty -Name 'CERT_CERTIFICATE_TEMPLATE' -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force

                }

            }      

        }
        until ($ColObj.Next() -eq -1)
        Clear-Variable -Name ColObj
        $certArr += $Cert      

    }

    $json = [pscustomobject]@{'data' = @($certArr)} | ConvertTo-Json
    [Console]::WriteLine($json)
 
}