package wendy;

public enum SheetTemplate {
    NAAM("naam", 1),
    FACTUUR_DATUM("Factuurdatum", 2),
    FACTUUR_NUMMER("Factuurnummer", 3),
    TOTAL_INCL_BTW("TOTAAL INCL. BTW", 4),
    BTW_21("BTW 21 %", 5),
    BTW_9("BTW 9 %", 6),
    TOTAAL_BTW("TOTAAL BTW", 7),
    VERZENDKOSTEN("VERZENDKOSTEN", 8),
    TE_BETALEN("TE BETALEN", 9);

    private String name;
    private Integer cell;

    SheetTemplate(String name, Integer cell) {
        this.name = name;
        this.cell = cell;
    }

    public String getName() {
        return name;
    }

    public Integer getCell() {
        return cell;
    }
}
