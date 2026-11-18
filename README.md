# Restoran satışları: interaktiv Excel dashboard

> Düsturlarla işləyən Excel dashboard (makro yoxdur, pivot cədvəl yoxdur): KPI kartları, iki açılan filtr və beş nativ qrafik. 1,000 restoran sifarişi əsasında.

[English](README.en.md) · **[Faylı yüklə](restaurant_sales_dashboard.xlsx)**

![Dashboard](images/dashboard_preview.png)

*Yuxarıdakı şəkil dashboard-un məzmununu göstərir (matplotlib ilə çəkilib). İnteraktiv versiya Excel faylıdır.*

## Nə edir

- **Dashboard səhifəsi:** KPI kartları (gəlir, sifariş sayı, orta çek, reytinq, satılan say, ən yaxşı kateqoriya), **Filial** və **Sifariş növü** açılan siyahıları, beş qrafik (aylıq gəlir, kateqoriya, filial, həftə günü, sifariş növü payı). Siyahıdan dəyər seçdikdə bütün KPI və qrafiklər yenilənir.
- **Analysis səhifəsi:** hər rəqəm canlı `SUMIFS` / `COUNTIFS` / `AVERAGEIFS` düsturudur, məntiq açıq və yoxlanıla biləndir.
- **Data səhifəsi:** 1,000 sifariş Excel cədvəli kimi, `Month` və `Weekday` köməkçi sütunları düsturla.
- **Notes səhifəsi:** mənbə, yoxlamalar, istifadə qaydası, məhdudiyyətlər.

## Əsas nəticələr

| Nəticə | Rəqəm |
|---|---|
| Gəlir / sifariş / orta çek | $25,799 / 1,000 / $25.80 |
| Əsas yeməklər biznesi daşıyır | sifarişlərin 47.5%-i, gəlirin 66%-i |
| Tələb üç günə cəmlənib | Cümə-Bazar: gəlirin 57%-i, sifarişlərin 58%-i |
| Ən yaxşı filial | Downtown Manhattan, $3,676 (14%) |
| Ən böyük kanal | Zalda yemək 42%; çatdırılma və götürmə çeki təxminən $1 böyükdür ($26.27 və $25.17) |
| Razılıq | reytinqlərin 71.5%-i 4 və ya 5; orta 3.92 |

## Yoxlamalar

Data yoxlamaları (`src/prepare_data.py`): unikal ID, boş yox, `Order Total = Quantity x Unit Price` hər sətirdə, reytinq 1-5. Düsturlar ayrıca hesablama mühərriki ilə yenidən hesablanıb və pandas ilə müqayisə edilib: Analysis səhifəsinin 121 rəqəmi üst-üstə düşür, süzülmüş ssenari (Miami Beach + Delivery) də. Bu mühitdə masaüstü Excel-də açıb baxa bilmədim.

## Məhdudiyyət

10 restoran üçün 2 ildə 1,000 sifariş nümunədir, aylıq dəyişmələr kiçik say səs-küyüdür. Filial və Top 10 sıralaması tam dataya görə sabitdir, filtrdə yenidən sıralanmır.

## Müəllif

**Alyaz Cabbarova**, data analitik olmaq istəyən mütəxəssis, Bakı.
[LinkedIn](https://www.linkedin.com/in/alyaz-cabbarova-450a6140b) · alyazcbbrv@gmail.com · [GitHub](https://github.com/alyazcbbrv-byte)

MIT lisenziyası.
