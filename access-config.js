/**
 * PIN-Zugang – bitte PINs für Ihre Umgebung anpassen (siehe access-config.example.js).
 */
window.MS365_ACCESS_CONFIG = {
    enabled: true,
    pins: ['MS365-Schule', 'IT-Team', '#kurtrocks','#KurtRocks!','HLAEbensee', 'HAK-Steyr'],
    adminPin: '#kurtrocksMS365',
    /** Betreiber-UPNs: sehen „Admin“ im Konto-Menü und dürfen Admin ohne Master-PIN öffnen */
    operatorUpns: ['kurt@kurtsoeser.at']
};
