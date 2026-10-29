import re
import sys

def test_i18n():
    with open('web_dashboard_direccion/index.html', 'r', encoding='utf-8') as f:
        html = f.read()

    ids = [
        'sidebarSubtitle', 'navGroupHeaderDg', 'navGroupHeaderCom', 'navGroupHeaderOps', 
        'navGroupHeaderRent', 'navGroupHeaderCtrl', 'navBtnInicioText', 'navBtnBiReunionText', 
        'navBtnBiSearchText', 'navBtnResumenText', 'navBtnTrackingMartaText', 'navBtnBiForecastText', 
        'navBtnBiChurnText', 'navBtnBiExcepcionesText', 'navBtnProduccionText', 'navBtnBiComplexityText', 
        'navBtnBiAtpText', 'navBtnRentabilidadCogsText', 'navBtnRentabilidadFullCostText', 
        'navBtnBiConcentracionText', 'navBtnControlesText', 'navBtnBiSnapshotsText', 'navBtnBiStagingText',
        'headerBtnLogout', 'tabBadgeControles'
    ]

    all_ok = True
    for i in ids:
        if f'id="{i}"' not in html:
            print(f'Missing ID: {i}')
            all_ok = False

    assert all_ok, "Some IDs missing in index.html"
    print("All 25 IDs exist in index.html!")

    # Check cockpit_inicio_logic.js syntax and translations
    with open('web_dashboard_direccion/cockpit_inicio_logic.js', 'r', encoding='utf-8') as f:
        cockpit_code = f.read()

    assert 'isEn' in cockpit_code, "isEn missing in cockpit_inicio_logic.js"
    assert 'Executive Management Cockpit' in cockpit_code, "English title missing"
    assert '4 Key Business Vital Signs' in cockpit_code, "English vital signs title missing"
    assert '1. Net Order Intake' in cockpit_code, "English card 1 missing"
    assert '2. Full Industrial Margin' in cockpit_code, "English card 2 missing"
    assert '3. Accounting vs Analytical EBT (Sep-26)' in cockpit_code, "English card 3 missing"
    assert '4. Pending Delivery Orders' in cockpit_code, "English card 4 missing"
    assert 'Matters Requiring Immediate Action (Ranked by € Impact)' in cockpit_code, "English radar title missing"
    print("All Cockpit English strings verified in cockpit_inicio_logic.js!")

if __name__ == '__main__':
    test_i18n()

