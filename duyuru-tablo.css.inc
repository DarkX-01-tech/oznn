    /* Duyuru içi tablolar - index.asp ile aynı görünüm */
    .duyuru-icerik table {
        width: 100%;
        max-width: 100%;
        border-collapse: collapse;
        margin: 12px 0;
        text-align: left;
        font-weight: 400;
        border: 1px solid #ddd;
        table-layout: auto;
    }

    .duyuru-icerik table td,
    .duyuru-icerik table th {
        border: 1px solid #ddd;
        padding: 8px 12px;
        text-align: left !important;
        font-weight: 400;
        vertical-align: top;
        word-break: break-word;
    }

    .duyuru-icerik table th {
        background-color: #f0f0f0;
        font-weight: 600;
        text-align: center !important;
    }

    .duyuru-icerik table tr:first-child td {
        background-color: #f0f0f0;
        font-weight: 600;
        text-align: center !important;
    }

    .duyuru-icerik table td:first-child {
        font-weight: 600;
        width: 40%;
    }

    .duyuru-icerik table b,
    .duyuru-icerik table strong {
        font-weight: 600 !important;
    }

    /* Tablolar gizleme alanından etkilenmesin */
    .image-content table,
    .duyuru-icerik > table {
        display: table !important;
        visibility: visible !important;
        opacity: 1 !important;
        max-height: none !important;
    }
