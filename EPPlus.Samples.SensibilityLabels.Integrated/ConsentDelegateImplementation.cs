using Microsoft.InformationProtection;
class ConsentDelegateImplementation : IConsentDelegate
{
    public Consent GetUserConsent(string url)
    {
        return Consent.Accept;
    }
}