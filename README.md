# Proje Açıklaması

Bu proje, Excel dosyalarında belirli koşullara göre hücreleri kontrol ederek veri manipülasyonu yapmak için Python ile yazılmıştır. Projenin amacı, çeşitli Excel dosyalarında hücrelerdeki içeriği analiz ederek, belirli koşullara uyan hücrelerin verilerini düzenlemek ve yeni çalışma sayfaları oluşturmak, gerektiğinde buton oluşturup vba kodu eklemek veya mevcut sayfaları güncellemektir.

## Gereksinimler

Bu proje için aşağıdaki gereksinimlerin karşılanması gerekmektedir:

- Python (versiyon 3.x)
- `win32com.client` kütüphanesi
- `os` kütüphanesi

## Nasıl Kullanılır

1. Projeyi çalıştırmadan önce, Python ve gerekli kütüphaneleri yüklediğinizden emin olun.

2. Proje klasöründe, Python betiği olan `Python-Excel_Duplicates_Sum.py` dosyasını çalıştırın.

3. Program başladığında, komut satırında talimatlar ve seçenekler belirecektir. İlgili talimatları izleyerek devam edin.

## Projenin İşleyişi

- Proje, `get_second_matching_cell()` adlı bir fonksiyon ile başlar. Bu fonksiyon, belirli koşullara uygun hücreleri bulup işler.
- Uygun hücrelerin belirlenmesinden sonra, ilgili veriler düzenlenir ve yeni çalışma sayfaları oluşturulur veya mevcut sayfalar güncellenir.
- Hücre içeriğine ve koşullara göre farklı işlemler yapılarak, Excel dosyalarındaki veriler düzenlenir ve yeni raporlar oluşturulur.

# Project Description

This project is written in Python to manipulate data by checking cells in Excel files based on specific conditions. The aim of the project is to analyze the contents of cells in various Excel files, modify data in cells that meet certain conditions, create new worksheets, add buttons and VBA code when necessary, and update existing worksheets.

## Requirements

To run this project, the following requirements must be met:

- Python (version 3.x)
- `win32com.client` library
- `os` library

## How to Use

1. Before running the project, make sure you have Python and the required libraries installed.

2. In the project folder, run the Python script named `Python-Excel_Duplicates_Sum.py`.

3. When the program starts, follow the instructions and options displayed in the command line.

## Project Workflow

- The project begins with the function `get_second_matching_cell()`, which finds and processes cells that meet certain conditions.
- After identifying the appropriate cells, relevant data is organized, and new worksheets are created or existing worksheets are updated.
- Different operations are performed based on cell content and conditions, resulting in data manipulation and the creation of new reports in Excel files.

