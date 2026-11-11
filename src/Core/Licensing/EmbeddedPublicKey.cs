namespace IMAP2M365.Licensing;

/// <summary>
/// Öffentlicher RSA-Schlüssel des Lizenzgebers (Rudolf). Nur dieser steckt im Tool.
/// Das Tool kann damit Lizenzen PRÜFEN und Anfragen VERSCHLÜSSELN, aber keine Lizenzen
/// fälschen – dafür wäre der private Schlüssel nötig (nur im Generator-Tool).
///
/// Für den Produktivbetrieb kannst du im Generator ein neues Schlüsselpaar erzeugen und
/// diesen PEM-Block ersetzen.
/// </summary>
public static class EmbeddedPublicKey
{
    public const string Pem = @"-----BEGIN PUBLIC KEY-----
MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIBigKCAYEA2z7jBw2m1OP6h1ZcCUWR
WH43VrUgbS/VRbhEk5qVlXv7XscS1lpnfOEaWkWnrlcWv2CDwxdwOo3YZhUC+6Yy
tpnxJDImCUFh9rbZ96Yz2APfFl2DCUWOhetzpvmvAN+sz7o8rdpajb9Sc0KToajI
msDldoMrNfr4E90I0WO1q1OPkseGGdzqfLi03yasLlXND34QftFaoY4ofFelSN5G
XuSQ2ggMNzrhPEVYYDmjsEBv7dx31aQej2+Gnc0+Ct61K5l/DgT9bceQDTGnuKOA
3BD/RMaE1yX3cCq/jf8mWG6x3MTNPxUmUAMptND0GK3JaH6OUH9HsbXCXtLfbf/i
PjDjNIfLach+55MrV/iXsS4yeAm1s5fJD41NlGfNa86p8JM+UECPEjm91xoqjrjU
3T2cHh5D7ucjtKLS51i7n0b5iel7lag2Dr7WqgUPyb88ALGv8wTsAEDX28ffyXN3
P4TJF+hV3OJGS2hUy1+8mg8Nc7k5yup9onKUgsfOwWEnAgMBAAE=
-----END PUBLIC KEY-----";
}
