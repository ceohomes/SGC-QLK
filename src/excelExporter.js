import * as XLSXStyleRaw from 'xlsx-js-style'
const XLSXStyle = XLSXStyleRaw.default || XLSXStyleRaw

export const exportConsolidatedExcel = (selectedRows, localProject, customCategoryMap, getUnitCategory, materialPriceRows, materialMetadataMap) => {
  if (!selectedRows || selectedRows.length === 0) return

  const wb = XLSXStyle.utils.book_new()

  const titleStyle = {
    font: { name: 'Segoe UI', sz: 16, bold: true, color: { rgb: '1E3A8A' } },
    alignment: { horizontal: 'left', vertical: 'center' }
  }
  const subtitleStyle = {
    font: { name: 'Segoe UI', sz: 11, italic: true, color: { rgb: '475569' } },
    alignment: { horizontal: 'left', vertical: 'center' }
  }
  
  const tblHeaderStyle1 = {
    fill: { patternType: 'solid', fgColor: { rgb: '0F766E' } },
    font: { name: 'Segoe UI', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'center', vertical: 'center' }
  }
  const tblHeaderStyle2 = {
    fill: { patternType: 'solid', fgColor: { rgb: 'C2410C' } },
    font: { name: 'Segoe UI', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'center', vertical: 'center' }
  }
  
  const thStyle1 = {
    fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
    font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top: { style: 'thin', color: { rgb: 'CCFBF1' } },
      bottom: { style: 'thin', color: { rgb: 'CCFBF1' } },
      left: { style: 'thin', color: { rgb: 'CCFBF1' } },
      right: { style: 'thin', color: { rgb: 'CCFBF1' } }
    }
  }
  const thStyle2 = {
    fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
    font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top: { style: 'thin', color: { rgb: 'FFEDD5' } },
      bottom: { style: 'thin', color: { rgb: 'FFEDD5' } },
      left: { style: 'thin', color: { rgb: 'FFEDD5' } },
      right: { style: 'thin', color: { rgb: 'FFEDD5' } }
    }
  }

  const dataCellStyle = {
    font: { name: 'Segoe UI', sz: 9.5, color: { rgb: '1E293B' } },
    border: {
      top: { style: 'thin', color: { rgb: 'E2E8F0' } },
      bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
      left: { style: 'thin', color: { rgb: 'E2E8F0' } },
      right: { style: 'thin', color: { rgb: 'E2E8F0' } }
    },
    alignment: { vertical: 'center', wrapText: true }
  }

  const legendLabelStyle = {
    font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
    alignment: { vertical: 'center' }
  }
  const legendNccStyle = {
    fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
    font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '78350F' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: 'FDE68A' } },
      bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
      left: { style: 'thin', color: { rgb: 'FDE68A' } },
      right: { style: 'thin', color: { rgb: 'FDE68A' } }
    }
  }
  const legendKhoStyle = {
    fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
    font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '1E3A8A' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: 'BFDBFE' } },
      bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
      left: { style: 'thin', color: { rgb: 'BFDBFE' } },
      right: { style: 'thin', color: { rgb: 'BFDBFE' } }
    }
  }
  const legendToDoiStyle = {
    fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
    font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '065F46' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: 'BBF7D0' } },
      bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
      left: { style: 'thin', color: { rgb: 'BBF7D0' } },
      right: { style: 'thin', color: { rgb: 'BBF7D0' } }
    }
  }

  const footerStyle = {
    font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E293B' } },
    fill: { patternType: 'solid', fgColor: { rgb: 'F8FAFC' } },
    border: {
      top: { style: 'thin', color: { rgb: '94A3B8' } },
      bottom: { style: 'double', color: { rgb: '94A3B8' } },
      left: { style: 'thin', color: { rgb: 'E2E8F0' } },
      right: { style: 'thin', color: { rgb: 'E2E8F0' } }
    },
    alignment: { vertical: 'center', wrapText: true }
  }

  const totalYellowStyle = {
    ...footerStyle,
    fill: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } },
    alignment: { horizontal: 'right', vertical: 'center', wrapText: true }
  }

  const sectionTitleStyle = {
    font: { name: 'Segoe UI', sz: 12, bold: true, color: { rgb: '1E3A8A' } },
    alignment: { horizontal: 'left', vertical: 'center' },
    fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } }
  }

  // Initialize 3 empty sheet structures
  const wsBaoCao = {}
  const wsThucNhap = {}
  const wsThucXuat = {}

  // Add initial general headers
  wsBaoCao['A1'] = { v: 'BẢNG TỔNG HỢP GIẢI TRÌNH CẤU THÀNH SỐ LIỆU', t: 's', s: titleStyle }
  wsBaoCao['A2'] = { v: `Kho chọn: ${localProject || 'Tất cả'} | Xuất gộp cho ${selectedRows.length} mặt hàng`, t: 's', s: subtitleStyle }

  wsThucNhap['A1'] = { v: 'BẢNG CHI TIẾT CHỨNG TỪ THỰC NHẬP (CẤU THÀNH SỐ LIỆU GỘP)', t: 's', s: titleStyle }
  wsThucNhap['A2'] = { v: `Kho chọn: ${localProject || 'Tất cả'} | Xuất gộp cho ${selectedRows.length} mặt hàng`, t: 's', s: subtitleStyle }

  wsThucXuat['A1'] = { v: 'BẢNG CHI TIẾT CHỨNG TỪ THỰC XUẤT (CẤU THÀNH SỐ LIỆU GỘP)', t: 's', s: titleStyle }
  wsThucXuat['A2'] = { v: `Kho chọn: ${localProject || 'Tất cả'} | Xuất gộp cho ${selectedRows.length} mặt hàng`, t: 's', s: subtitleStyle }

  wsBaoCao['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 16 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 16 } }
  ]
  wsThucNhap['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 11 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 11 } }
  ]
  wsThucXuat['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 11 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 11 } }
  ]

  let rowBaoCao = 4
  let rowThucNhap = 4
  let rowThucXuat = 4

  selectedRows.forEach(itemRow => {
    const maSAP = itemRow.maSAP
    const tenVatTu = itemRow.tenVatTu
    const dvt = itemRow.dvt

    const txs = itemRow.transactions || []
    const nhapSummaryByUnit = {}
    const xuatSummaryByUnit = {}

    txs.forEach(tx => {
      const normProject = String(localProject || '').trim().toLowerCase()
      const isGiaoProject = String(tx.donViGiao || '').trim().toLowerCase() === normProject
      const isNhanProject = String(tx.donViNhan || '').trim().toLowerCase() === normProject

      if (tx.nhapVal > 0) {
        let role = 'Đơn vị nhận'
        let partner = tx.donViGiao || 'Chưa xác định'
        if (isGiaoProject) {
          role = 'Đơn vị giao'
          partner = tx.donViNhan || 'Chưa xác định'
        }

        const key = `${partner}||${tx.detailTypeNhap}||${role}`
        if (!nhapSummaryByUnit[key]) {
          nhapSummaryByUnit[key] = {
            partner,
            role,
            detailType: tx.detailTypeNhap,
            logicVal: tx.logicNhapVal,
            explain: tx.explainNhap,
            totalQty: 0,
            totalContribution: 0
          }
        }
        nhapSummaryByUnit[key].totalQty += tx.nhapVal
        nhapSummaryByUnit[key].totalContribution += tx.nhapVal * tx.logicNhapVal
      }
      if (tx.xuatVal > 0) {
        let role = 'Đơn vị giao'
        let partner = tx.donViNhan || 'Chưa xác định'
        if (isNhanProject) {
          role = 'Đơn vị nhận'
          partner = tx.donViGiao || 'Chưa xác định'
        }

        const key = `${partner}||${tx.detailTypeXuat}||${role}`
        if (!xuatSummaryByUnit[key]) {
          xuatSummaryByUnit[key] = {
            partner,
            role,
            detailType: tx.detailTypeXuat,
            logicVal: tx.logicXuatVal,
            explain: tx.explainXuat,
            totalQty: 0,
            totalContribution: 0
          }
        }
        xuatSummaryByUnit[key].totalQty += tx.xuatVal
        xuatSummaryByUnit[key].totalContribution += tx.xuatVal * tx.logicXuatVal
      }
    })

    const sortNhapSummaryList = (list) => {
      return [...list].sort((a, b) => {
        const rolePriorityA = a.role === 'Đơn vị nhận' ? 1 : a.role === 'Đơn vị giao' ? 2 : 3
        const rolePriorityB = b.role === 'Đơn vị nhận' ? 1 : b.role === 'Đơn vị giao' ? 2 : 3
        if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

        const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
        const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
        const catPriorityA = catA === 'ncc' ? 1 : catA === 'kho' ? 2 : catA === 'todoi' ? 3 : 4

        const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
        const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
        const catPriorityB = catB === 'ncc' ? 1 : catB === 'kho' ? 2 : catB === 'todoi' ? 3 : 4

        if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

        const valA = Math.abs(a.totalContribution || 0)
        const valB = Math.abs(b.totalContribution || 0)
        if (valB !== valA) return valB - valA

        return (a.partner || '').localeCompare(b.partner || '', 'vi')
      })
    }

    const sortXuatSummaryList = (list) => {
      return [...list].sort((a, b) => {
        const rolePriorityA = a.role === 'Đơn vị giao' ? 1 : a.role === 'Đơn vị nhận' ? 2 : 3
        const rolePriorityB = b.role === 'Đơn vị giao' ? 1 : b.role === 'Đơn vị nhận' ? 2 : 3
        if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

        const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
        const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
        const catPriorityA = catA === 'todoi' ? 1 : catA === 'kho' ? 2 : catA === 'ncc' ? 3 : 4

        const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
        const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
        const catPriorityB = catB === 'todoi' ? 1 : catB === 'kho' ? 2 : catB === 'ncc' ? 3 : 4

        if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

        const valA = Math.abs(a.totalContribution || 0)
        const valB = Math.abs(b.totalContribution || 0)
        if (valB !== valA) return valB - valA

        return (a.partner || '').localeCompare(b.partner || '', 'vi')
      })
    }

    const nhapSummaryList = sortNhapSummaryList(Object.values(nhapSummaryByUnit))
    const xuatSummaryList = sortXuatSummaryList(Object.values(xuatSummaryByUnit))

    const groupAndSortByPartnerForMulti = (list) => {
      const grouped = {}
      list.forEach(item => {
        const normPartner = (item.partner || '').trim()
        if (!grouped[normPartner]) {
          grouped[normPartner] = {
            partner: normPartner,
            role: item.role,
            detailType: item.detailType,
            logicVal: item.logicVal,
            explain: item.explain,
            totalQty: 0,
            totalContribution: 0,
            items: []
          }
        }
        grouped[normPartner].totalQty += item.totalQty
        grouped[normPartner].totalContribution += item.totalQty * item.logicVal
        grouped[normPartner].items.push(item)
      })

      const groupedList = Object.values(grouped).map(group => {
        group.items.sort((a, b) => b.totalQty - a.totalQty)
        const dominant = group.items[0]
        return {
          partner: group.partner,
          role: dominant.role,
          detailType: dominant.detailType,
          logicVal: dominant.logicVal,
          explain: dominant.explain || '',
          totalQty: group.totalQty,
          totalContribution: group.totalContribution
        }
      })

      groupedList.sort((a, b) => b.totalContribution - a.totalContribution)
      return groupedList
    }

    const nhapListBaoCao = nhapSummaryList.filter(item => item.logicVal !== 0)
    const xuatListBaoCao = xuatSummaryList.filter(item => item.logicVal !== 0)

    const groupedNhapBaoCao = groupAndSortByPartnerForMulti(nhapListBaoCao)
    groupedNhapBaoCao.sort((a, b) => {
      const rolePriorityA = a.role === 'Đơn vị nhận' ? 1 : a.role === 'Đơn vị giao' ? 2 : 3
      const rolePriorityB = b.role === 'Đơn vị nhận' ? 1 : b.role === 'Đơn vị giao' ? 2 : 3
      if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

      const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
      const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
      const catPriorityA = catA === 'ncc' ? 1 : catA === 'kho' ? 2 : catA === 'todoi' ? 3 : 4

      const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
      const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
      const catPriorityB = catB === 'ncc' ? 1 : catB === 'kho' ? 2 : catB === 'todoi' ? 3 : 4

      if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

      const valA = Math.abs(a.totalContribution || 0)
      const valB = Math.abs(b.totalContribution || 0)
      if (valB !== valA) return valB - valA

      return (a.partner || '').localeCompare(b.partner || '', 'vi')
    })

    const groupedXuatBaoCao = groupAndSortByPartnerForMulti(xuatListBaoCao)
    groupedXuatBaoCao.sort((a, b) => {
      const rolePriorityA = a.role === 'Đơn vị giao' ? 1 : a.role === 'Đơn vị nhận' ? 2 : 3
      const rolePriorityB = b.role === 'Đơn vị giao' ? 1 : b.role === 'Đơn vị nhận' ? 2 : 3
      if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

      const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
      const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
      const catPriorityA = catA === 'todoi' ? 1 : catA === 'kho' ? 2 : catA === 'ncc' ? 3 : 4

      const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
      const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
      const catPriorityB = catB === 'todoi' ? 1 : catB === 'kho' ? 2 : catB === 'ncc' ? 3 : 4

      if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

      const valA = Math.abs(a.totalContribution || 0)
      const valB = Math.abs(b.totalContribution || 0)
      if (valB !== valA) return valB - valA

      return (a.partner || '').localeCompare(b.partner || '', 'vi')
    })

    const sheetMaxRows = Math.max(groupedNhapBaoCao.length, groupedXuatBaoCao.length)

    // --- SECTION HEADER ON SHEET 1 (BÁO CÁO GIẢI TRÌNH) ---
    wsBaoCao['!merges'].push({ s: { r: rowBaoCao - 1, c: 0 }, e: { r: rowBaoCao - 1, c: 16 } })
    wsBaoCao[`A${rowBaoCao}`] = { v: `MẶT HÀNG: ${tenVatTu.toUpperCase()} (${maSAP}) | ĐVT: ${dvt}`, t: 's', s: sectionTitleStyle }

    // Legend row
    wsBaoCao[`A${rowBaoCao + 1}`] = { v: 'Màu sắc đối tác:', t: 's', s: legendLabelStyle }
    wsBaoCao[`C${rowBaoCao + 1}`] = { v: '■ Nhà cung cấp (NCC)', t: 's', s: legendNccStyle }
    wsBaoCao[`E${rowBaoCao + 1}`] = { v: '■ Kho khác', t: 's', s: legendKhoStyle }
    wsBaoCao[`G${rowBaoCao + 1}`] = { v: '■ Tổ đội', t: 's', s: legendToDoiStyle }

    // Table titles
    wsBaoCao['!merges'].push({ s: { r: rowBaoCao + 2, c: 0 }, e: { r: rowBaoCao + 2, c: 7 } })
    wsBaoCao['!merges'].push({ s: { r: rowBaoCao + 2, c: 9 }, e: { r: rowBaoCao + 2, c: 16 } })
    wsBaoCao[`A${rowBaoCao + 3}`] = { v: 'BẢNG TỔNG HỢP DIỄN GIẢI THỰC NHẬP (SUMIFS THEO ĐƠN VỊ)', t: 's', s: tblHeaderStyle1 }
    wsBaoCao[`J${rowBaoCao + 3}`] = { v: 'BẢNG TỔNG HỢP DIỄN GIẢI THỰC XUẤT (SUMIFS THEO ĐƠN VỊ)', t: 's', s: tblHeaderStyle2 }

    // Column subheaders
    const sub1 = ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Thực nhập', 'Ghi chú']
    const sub2 = ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Thực xuất', 'Ghi chú']

    sub1.forEach((label, i) => {
      const colChar = String.fromCharCode(65 + i)
      wsBaoCao[`${colChar}${rowBaoCao + 4}`] = { v: label, t: 's', s: thStyle1 }
    })

    sub2.forEach((label, i) => {
      const colChar = String.fromCharCode(74 + i)
      wsBaoCao[`${colChar}${rowBaoCao + 4}`] = { v: label, t: 's', s: thStyle2 }
    })

    // Section total row with dynamic formulas
    wsBaoCao[`A${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`B${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`C${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`D${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`E${rowBaoCao + 5}`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsBaoCao[`F${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`G${rowBaoCao + 5}`] = { f: sheetMaxRows > 0 ? `SUM(G${rowBaoCao + 6}:G${rowBaoCao + 5 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
    wsBaoCao[`H${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }

    wsBaoCao[`I${rowBaoCao + 5}`] = { v: '', t: 's', s: { font: { name: 'Segoe UI', sz: 9.5 } } }

    wsBaoCao[`J${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`K${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`L${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`M${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`N${rowBaoCao + 5}`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsBaoCao[`O${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }
    wsBaoCao[`P${rowBaoCao + 5}`] = { f: sheetMaxRows > 0 ? `SUM(P${rowBaoCao + 6}:P${rowBaoCao + 5 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
    wsBaoCao[`Q${rowBaoCao + 5}`] = { v: '', t: 's', s: footerStyle }

    // Section data rows on Sheet 1
    for (let i = 0; i < sheetMaxRows; i++) {
      const rowNum = rowBaoCao + 6 + i

      if (i < groupedNhapBaoCao.length) {
        const item = groupedNhapBaoCao[i]
        const labelType = item.detailType === 'nhan_ncc' ? 'Nhận từ NCC'
                        : item.detailType === 'nhan_kho' ? 'Nhận từ Kho gửi'
                        : item.detailType === 'tra_ncc' ? 'Trả lại NCC'
                        : item.detailType === 'dieu_chuyen_di' ? 'Điều chuyển đi (Không tính)'
                        : item.detailType === 'nhan_khac' ? 'Nhận nguồn khác (Không tính)'
                        : 'Không tính'

        const norm = (item.partner || '').trim().replace(/\s+/g, ' ')
        const cat = (customCategoryMap && customCategoryMap[norm]) || getUnitCategory(norm)

        let partnerStyle = { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '334155' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } }
        if (cat === 'kho') {
          partnerStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
          }
        } else if (cat === 'ncc') {
          partnerStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
          }
        } else if (cat === 'todoi') {
          partnerStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
          }
        }

        const isZero = item.logicVal === 0
        const isNegative = item.logicVal < 0
        let labelStyle = { ...dataCellStyle }
        if (isZero) {
          labelStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'CBD5E1' } },
              bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
              left: { style: 'thin', color: { rgb: 'CBD5E1' } },
              right: { style: 'thin', color: { rgb: 'CBD5E1' } }
            }
          }
        } else if (isNegative) {
          labelStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'FCA5A5' } },
              bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
              left: { style: 'thin', color: { rgb: 'FCA5A5' } },
              right: { style: 'thin', color: { rgb: 'FCA5A5' } }
            }
          }
        } else {
          labelStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'A7F3D0' } },
              bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
              left: { style: 'thin', color: { rgb: 'A7F3D0' } },
              right: { style: 'thin', color: { rgb: 'A7F3D0' } }
            }
          }
        }

        let roleStyle = { ...dataCellStyle }
        if (item.role === 'Đơn vị nhận') {
          roleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '6D28D9' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F5F3FF' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'DDD6FE' } },
              bottom: { style: 'thin', color: { rgb: 'DDD6FE' } },
              left: { style: 'thin', color: { rgb: 'DDD6FE' } },
              right: { style: 'thin', color: { rgb: 'DDD6FE' } }
            }
          }
        } else {
          roleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'EA580C' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'FED7AA' } },
              bottom: { style: 'thin', color: { rgb: 'FED7AA' } },
              left: { style: 'thin', color: { rgb: 'FED7AA' } },
              right: { style: 'thin', color: { rgb: 'FED7AA' } }
            }
          }
        }

        const khoStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: roleStyle.font.color },
          fill: roleStyle.fill,
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: roleStyle.border
        }

        const catLabel = cat === 'ncc' ? 'Nhà cung cấp (NCC)'
                       : cat === 'kho' ? 'Kho khác'
                       : cat === 'todoi' ? 'Tổ đội'
                       : ''
        let partnerRoleStyle = { ...dataCellStyle }
        if (cat === 'ncc') {
          partnerRoleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'FDE68A' } },
              bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
              left: { style: 'thin', color: { rgb: 'FDE68A' } },
              right: { style: 'thin', color: { rgb: 'FDE68A' } }
            }
          }
        } else if (cat === 'kho') {
          partnerRoleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'BFDBFE' } },
              bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
              left: { style: 'thin', color: { rgb: 'BFDBFE' } },
              right: { style: 'thin', color: { rgb: 'BFDBFE' } }
            }
          }
        } else if (cat === 'todoi') {
          partnerRoleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'BBF7D0' } },
              bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
              left: { style: 'thin', color: { rgb: 'BBF7D0' } },
              right: { style: 'thin', color: { rgb: 'BBF7D0' } }
            }
          }
        }

        wsBaoCao[`A${rowNum}`] = { v: i + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
        wsBaoCao[`B${rowNum}`] = { v: labelType, t: 's', s: labelStyle }
        wsBaoCao[`C${rowNum}`] = { v: localProject || 'Tất cả', t: 's', s: khoStyle }
        wsBaoCao[`D${rowNum}`] = { v: item.role || 'Đơn vị nhận', t: 's', s: roleStyle }
        wsBaoCao[`E${rowNum}`] = { v: item.partner, t: 's', s: partnerStyle }
        wsBaoCao[`F${rowNum}`] = { v: catLabel, t: 's', s: partnerRoleStyle }
        
        const partnerCol = item.role === 'Đơn vị giao' ? 'G' : 'F'
        wsBaoCao[`G${rowNum}`] = { 
          f: `SUMIFS('Chi tiết Thực nhập'!L:L, 'Chi tiết Thực nhập'!${partnerCol}:${partnerCol}, E${rowNum}, 'Chi tiết Thực nhập'!B:B, "${maSAP}")`, 
          t: 'n', 
          z: '#,##0.00;[Red](#,##0.00);"-"',
          s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } } } 
        }
        wsBaoCao[`H${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
      } else {
        ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(col => {
          wsBaoCao[`${col}${rowNum}`] = { v: '', t: 's', s: { ...dataCellStyle, border: {} } }
        })
      }

      wsBaoCao[`I${rowNum}`] = { v: '', t: 's', s: { font: { name: 'Segoe UI', sz: 9.5 } } }

      if (i < groupedXuatBaoCao.length) {
        const item = groupedXuatBaoCao[i]
        const labelType = item.detailType === 'xuat_todoi' ? 'Xuất cho Tổ đội'
                        : item.detailType === 'xuat_kho' ? 'Xuất đi Kho nhận'
                        : item.detailType === 'todoi_tra' ? 'Tổ đội trả hàng'
                        : item.detailType === 'xuat_khac' ? 'Xuất khác (Không tính)'
                        : item.detailType === 'nhan_lai_khac' ? 'Nhận lại khác (Không tính)'
                        : 'Không tính'

        const norm = (item.partner || '').trim().replace(/\s+/g, ' ')
        const cat = (customCategoryMap && customCategoryMap[norm]) || getUnitCategory(norm)

        let partnerStyle = { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '334155' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } }
        if (cat === 'kho') {
          partnerStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
          }
        } else if (cat === 'ncc') {
          partnerStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
          }
        } else if (cat === 'todoi') {
          partnerStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
          }
        }

        const isZero = item.logicVal === 0
        const isNegative = item.logicVal < 0
        let labelStyle = { ...dataCellStyle }
        if (isZero) {
          labelStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'CBD5E1' } },
              bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
              left: { style: 'thin', color: { rgb: 'CBD5E1' } },
              right: { style: 'thin', color: { rgb: 'CBD5E1' } }
            }
          }
        } else if (isNegative) {
          labelStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'FCA5A5' } },
              bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
              left: { style: 'thin', color: { rgb: 'FCA5A5' } },
              right: { style: 'thin', color: { rgb: 'FCA5A5' } }
            }
          }
        } else {
          labelStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'A7F3D0' } },
              bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
              left: { style: 'thin', color: { rgb: 'A7F3D0' } },
              right: { style: 'thin', color: { rgb: 'A7F3D0' } }
            }
          }
        }

        let roleStyle = { ...dataCellStyle }
        if (item.role === 'Đơn vị nhận') {
          roleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '6D28D9' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F5F3FF' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'DDD6FE' } },
              bottom: { style: 'thin', color: { rgb: 'DDD6FE' } },
              left: { style: 'thin', color: { rgb: 'DDD6FE' } },
              right: { style: 'thin', color: { rgb: 'DDD6FE' } }
            }
          }
        } else {
          roleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'EA580C' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'FED7AA' } },
              bottom: { style: 'thin', color: { rgb: 'FED7AA' } },
              left: { style: 'thin', color: { rgb: 'FED7AA' } },
              right: { style: 'thin', color: { rgb: 'FED7AA' } }
            }
          }
        }

        const khoStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: roleStyle.font.color },
          fill: roleStyle.fill,
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: roleStyle.border
        }

        const catLabel = cat === 'ncc' ? 'Nhà cung cấp (NCC)'
                       : cat === 'kho' ? 'Kho khác'
                       : cat === 'todoi' ? 'Tổ đội'
                       : ''
        let partnerRoleStyle = { ...dataCellStyle }
        if (cat === 'ncc') {
          partnerRoleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'FDE68A' } },
              bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
              left: { style: 'thin', color: { rgb: 'FDE68A' } },
              right: { style: 'thin', color: { rgb: 'FDE68A' } }
            }
          }
        } else if (cat === 'kho') {
          partnerRoleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'BFDBFE' } },
              bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
              left: { style: 'thin', color: { rgb: 'BFDBFE' } },
              right: { style: 'thin', color: { rgb: 'BFDBFE' } }
            }
          }
        } else if (cat === 'todoi') {
          partnerRoleStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'BBF7D0' } },
              bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
              left: { style: 'thin', color: { rgb: 'BBF7D0' } },
              right: { style: 'thin', color: { rgb: 'BBF7D0' } }
            }
          }
        }

        wsBaoCao[`J${rowNum}`] = { v: i + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
        wsBaoCao[`K${rowNum}`] = { v: labelType, t: 's', s: labelStyle }
        wsBaoCao[`L${rowNum}`] = { v: localProject || 'Tất cả', t: 's', s: khoStyle }
        wsBaoCao[`M${rowNum}`] = { v: item.role || 'Đơn vị giao', t: 's', s: roleStyle }
        wsBaoCao[`N${rowNum}`] = { v: item.partner, t: 's', s: partnerStyle }
        wsBaoCao[`O${rowNum}`] = { v: catLabel, t: 's', s: partnerRoleStyle }
        
        const partnerColXuat = item.role === 'Đơn vị giao' ? 'G' : 'F'
        wsBaoCao[`P${rowNum}`] = { 
          f: `SUMIFS('Chi tiết Thực xuất'!L:L, 'Chi tiết Thực xuất'!${partnerColXuat}:${partnerColXuat}, N${rowNum}, 'Chi tiết Thực xuất'!B:B, "${maSAP}")`, 
          t: 'n', 
          z: '#,##0.00;[Red](#,##0.00);"-"',
          s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } } } 
        }
        wsBaoCao[`Q${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
      } else {
        ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'].forEach(col => {
          wsBaoCao[`${col}${rowNum}`] = { v: '', t: 's', s: { ...dataCellStyle, border: {} } }
        })
      }
    }

    // Advance Sheet 1 row pointer
    rowBaoCao += 6 + sheetMaxRows + 4


    // --- SECTION HEADER ON SHEET 2 (CHI TIẾT THỰC NHẬP) ---
    wsThucNhap['!merges'].push({ s: { r: rowThucNhap - 1, c: 0 }, e: { r: rowThucNhap - 1, c: 11 } })
    wsThucNhap[`A${rowThucNhap}`] = { v: `DANH SÁCH CHỨNG TỪ THỰC NHẬP - VẬT TƯ: ${tenVatTu.toUpperCase()} (${maSAP})`, t: 's', s: sectionTitleStyle }

    const s2Headers = ['STT', 'Mã SAP', 'Tên vật tư', 'Ngày', 'Mã chứng từ', 'Đơn vị giao (NCC/Kho)', 'Đơn vị nhận (Kho nhận)', 'KL Chứng từ', 'Phân loại đơn vị', 'Hệ số logic', 'Quy tắc & Diễn giải logic', 'Đóng góp thực']
    s2Headers.forEach((label, i) => {
      const colChar = String.fromCharCode(65 + i)
      wsThucNhap[`${colChar}${rowThucNhap + 1}`] = { v: label, t: 's', s: thStyle1 }
    })

    const nhapTxList = txs.filter(tx => tx.nhapVal > 0)
    const footerStyle1 = { ...footerStyle, fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } } }

    // Total row
    wsThucNhap[`A${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`B${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`C${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`D${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`E${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`F${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`G${rowThucNhap + 2}`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsThucNhap[`H${rowThucNhap + 2}`] = { f: nhapTxList.length > 0 ? `SUM(H${rowThucNhap + 3}:H${rowThucNhap + 2 + nhapTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsThucNhap[`I${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`J${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`K${rowThucNhap + 2}`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`L${rowThucNhap + 2}`] = { f: nhapTxList.length > 0 ? `SUM(L${rowThucNhap + 3}:L${rowThucNhap + 2 + nhapTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }

    // Data rows for Sheet 2
    nhapTxList.forEach((tx, idx) => {
      const rowNum = rowThucNhap + 3 + idx
      const typeLabel = tx.detailTypeNhap === 'nhan_ncc' ? 'Nhận từ NCC'
                      : tx.detailTypeNhap === 'nhan_kho' ? 'Nhận từ Kho gửi'
                      : tx.detailTypeNhap === 'tra_ncc' ? 'Trả lại NCC'
                      : tx.detailTypeNhap === 'dieu_chuyen_di' ? 'Điều chuyển đi (Không tính)'
                      : tx.detailTypeNhap === 'nhan_khac' ? 'Nhận nguồn khác (Không tính)'
                      : 'Không tính'

      const isZero = tx.logicNhapVal === 0
      const isNegative = tx.logicNhapVal < 0
      let typeStyle = { ...dataCellStyle }
      if (isZero) {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'CBD5E1' } },
            bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        }
      } else if (isNegative) {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'FCA5A5' } },
            bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
            left: { style: 'thin', color: { rgb: 'FCA5A5' } },
            right: { style: 'thin', color: { rgb: 'FCA5A5' } }
          }
        }
      } else {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'A7F3D0' } },
            bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
            left: { style: 'thin', color: { rgb: 'A7F3D0' } },
            right: { style: 'thin', color: { rgb: 'A7F3D0' } }
          }
        }
      }

      wsThucNhap[`A${rowNum}`] = { v: idx + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
      wsThucNhap[`B${rowNum}`] = { v: maSAP, t: 's', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true } } }
      wsThucNhap[`C${rowNum}`] = { v: tenVatTu, t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucNhap[`D${rowNum}`] = { v: tx.ngayXuatNhap || '', t: 's', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
      wsThucNhap[`E${rowNum}`] = { v: tx.maDonNhapKho || tx.maDonXuatKho || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true }, alignment: { vertical: 'center', wrapText: true } } }
      wsThucNhap[`F${rowNum}`] = { v: tx.donViGiao || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucNhap[`G${rowNum}`] = { v: tx.donViNhan || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucNhap[`H${rowNum}`] = { v: tx.nhapVal, t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
      wsThucNhap[`I${rowNum}`] = { v: typeLabel, t: 's', s: typeStyle }
      wsThucNhap[`J${rowNum}`] = { 
        v: tx.logicNhapVal, 
        t: 'n', 
        z: '#,##0;(#,##0);"-"', 
        s: { 
          ...dataCellStyle, 
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
          font: { 
            name: 'Segoe UI', 
            sz: 9.5, 
            bold: true,
            color: tx.logicNhapVal === -1 ? { rgb: 'FF0000' } : undefined
          } 
        } 
      }
      wsThucNhap[`K${rowNum}`] = { v: tx.explainNhap || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucNhap[`L${rowNum}`] = { 
        f: `H${rowNum} * J${rowNum}`, 
        t: 'n', 
        z: '#,##0.00;[Red](#,##0.00);"-"',
        s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } } } 
      }
    })

    // Advance Sheet 2 row pointer
    rowThucNhap += 3 + nhapTxList.length + 4


    // --- SECTION HEADER ON SHEET 3 (CHI TIẾT THỰC XUẤT) ---
    wsThucXuat['!merges'].push({ s: { r: rowThucXuat - 1, c: 0 }, e: { r: rowThucXuat - 1, c: 11 } })
    wsThucXuat[`A${rowThucXuat}`] = { v: `DANH SÁCH CHỨNG TỪ THỰC XUẤT - VẬT TƯ: ${tenVatTu.toUpperCase()} (${maSAP})`, t: 's', s: sectionTitleStyle }

    const s3Headers = ['STT', 'Mã SAP', 'Tên vật tư', 'Ngày', 'Mã chứng từ', 'Đơn vị giao (Kho gửi)', 'Đơn vị nhận (Tổ đội/Kho nhận)', 'KL Chứng từ', 'Phân loại đơn vị', 'Hệ số logic', 'Quy tắc & Diễn giải logic', 'Đóng góp thực']
    s3Headers.forEach((label, i) => {
      const colChar = String.fromCharCode(65 + i)
      wsThucXuat[`${colChar}${rowThucXuat + 1}`] = { v: label, t: 's', s: thStyle2 }
    })

    const xuatTxList = txs.filter(tx => tx.xuatVal > 0)
    const footerStyle2 = { ...footerStyle, fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } } }

    // Total row
    wsThucXuat[`A${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`B${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`C${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`D${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`E${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`F${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`G${rowThucXuat + 2}`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsThucXuat[`H${rowThucXuat + 2}`] = { f: xuatTxList.length > 0 ? `SUM(H${rowThucXuat + 3}:H${rowThucXuat + 2 + xuatTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsThucXuat[`I${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`J${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`K${rowThucXuat + 2}`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`L${rowThucXuat + 2}`] = { f: xuatTxList.length > 0 ? `SUM(L${rowThucXuat + 3}:L${rowThucXuat + 2 + xuatTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }

    // Data rows for Sheet 3
    xuatTxList.forEach((tx, idx) => {
      const rowNum = rowThucXuat + 3 + idx
      const typeLabel = tx.detailTypeXuat === 'xuat_todoi' ? 'Xuất cho Tổ đội'
                      : tx.detailTypeXuat === 'xuat_kho' ? 'Xuất đi Kho nhận'
                      : tx.detailTypeXuat === 'todoi_tra' ? 'Tổ đội trả hàng'
                      : tx.detailTypeXuat === 'xuat_khac' ? 'Xuất khác (Không tính)'
                      : tx.detailTypeXuat === 'nhan_lai_khac' ? 'Nhận lại khác (Không tính)'
                      : 'Không tính'

      const isZero = tx.logicXuatVal === 0
      const isNegative = tx.logicXuatVal < 0
      let typeStyle = { ...dataCellStyle }
      if (isZero) {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'CBD5E1' } },
            bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        }
      } else if (isNegative) {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'FCA5A5' } },
            bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
            left: { style: 'thin', color: { rgb: 'FCA5A5' } },
            right: { style: 'thin', color: { rgb: 'FCA5A5' } }
          }
        }
      } else {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'A7F3D0' } },
            bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
            left: { style: 'thin', color: { rgb: 'A7F3D0' } },
            right: { style: 'thin', color: { rgb: 'A7F3D0' } }
          }
        }
      }

      wsThucXuat[`A${rowNum}`] = { v: idx + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
      wsThucXuat[`B${rowNum}`] = { v: maSAP, t: 's', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true } } }
      wsThucXuat[`C${rowNum}`] = { v: tenVatTu, t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucXuat[`D${rowNum}`] = { v: tx.ngayXuatNhap || '', t: 's', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
      wsThucXuat[`E${rowNum}`] = { v: tx.maDonNhapKho || tx.maDonXuatKho || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true }, alignment: { vertical: 'center', wrapText: true } } }
      wsThucXuat[`F${rowNum}`] = { v: tx.donViGiao || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucXuat[`G${rowNum}`] = { v: tx.donViNhan || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucXuat[`H${rowNum}`] = { v: tx.xuatVal, t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
      wsThucXuat[`I${rowNum}`] = { v: typeLabel, t: 's', s: typeStyle }
      wsThucXuat[`J${rowNum}`] = { 
        v: tx.logicXuatVal, 
        t: 'n', 
        z: '#,##0;(#,##0);"-"', 
        s: { 
          ...dataCellStyle, 
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
          font: { 
            name: 'Segoe UI', 
            sz: 9.5, 
            bold: true,
            color: tx.logicXuatVal === -1 ? { rgb: 'FF0000' } : undefined
          } 
        } 
      }
      wsThucXuat[`K${rowNum}`] = { v: tx.explainXuat || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucXuat[`L${rowNum}`] = { 
        f: `H${rowNum} * J${rowNum}`, 
        t: 'n', 
        z: '#,##0.00;[Red](#,##0.00);"-"',
        s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } } } 
      }
    })

    // Advance Sheet 3 row pointer
    rowThucXuat += 3 + xuatTxList.length + 4
  })

  // Set Sheet 1 column widths
  wsBaoCao['!cols'] = [
    { wpx: 40 },  // A: STT
    { wpx: 130 }, // B: Phân loại đơn vị
    { wpx: 180 }, // C: Kho chọn
    { wch: 16 },  // D: Vai trò Kho chọn
    { wpx: 250 }, // E: Tên NCC/ Kho khác / Tổ đội
    { wch: 16 },  // F: Vai trò NCC/ Kho khác/ Tổ đội
    { wch: 16 },  // G: Thực nhập
    { wpx: 200 }, // H: Ghi chú
    { wpx: 40 },  // I: separator
    { wpx: 40 },  // J: STT
    { wpx: 130 }, // K: Phân loại đơn vị
    { wpx: 180 }, // L: Kho chọn
    { wch: 16 },  // M: Vai trò Kho chọn
    { wpx: 250 }, // N: Tên NCC/ Kho khác / Tổ đội
    { wch: 16 },  // O: Vai trò NCC/ Kho khác/ Tổ đội
    { wch: 16 },  // P: Thực xuất
    { wpx: 200 }  // Q: Ghi chú
  ]

  // Set Sheet 2 & 3 column widths
  const detailCols = [
    { wpx: 50 },  // A: STT
    { wpx: 100 }, // B: Mã SAP
    { wpx: 180 }, // C: Tên vật tư
    { wpx: 100 }, // D: Ngày
    { wpx: 130 }, // E: Mã chứng từ
    { wpx: 240 }, // F: Đơn vị giao
    { wpx: 240 }, // G: Đơn vị nhận
    { wpx: 120 }, // H: KL Chứng từ
    { wpx: 150 }, // I: Phân loại đơn vị
    { wpx: 80 },  // J: Hệ số logic
    { wpx: 250 }, // K: Quy tắc & Diễn giải logic
    { wpx: 120 }  // L: Đóng góp thực
  ]
  wsThucNhap['!cols'] = detailCols
  wsThucXuat['!cols'] = detailCols

  wsBaoCao['!ref'] = `A1:Q${rowBaoCao - 1}`
  wsThucNhap['!ref'] = `A1:L${rowThucNhap - 1}`
  wsThucXuat['!ref'] = `A1:L${rowThucXuat - 1}`

  // Append sheets to workbook
  XLSXStyle.utils.book_append_sheet(wb, wsBaoCao, "Báo cáo")
  XLSXStyle.utils.book_append_sheet(wb, wsThucNhap, "Chi tiết Thực nhập")
  XLSXStyle.utils.book_append_sheet(wb, wsThucXuat, "Chi tiết Thực xuất")

  if (materialPriceRows && materialPriceRows.length > 0) {
    const phanNhomSheet = buildPhanNhomVatTuSheet(materialPriceRows, materialMetadataMap)
    XLSXStyle.utils.book_append_sheet(wb, phanNhomSheet, "Phân nhóm Vật tư")
  }

  const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length)
    const view = new Uint8Array(buf)
    for (let j = 0; j < s.length; j++) view[j] = s.charCodeAt(j) & 0xFF
    return buf
  }

  const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `Giai_Trinh_Gop_${selectedRows.length}_Vat_Tu_${new Date().toISOString().slice(0, 10)}.xlsx`
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

export const buildPhanNhomVatTuSheet = (materialPriceRows, materialMetadataMap) => {
  const pws = {}
  const pcols = [
    { key: 'STT', label: 'STT', width: 50 },
    { key: 'maSAP', label: 'Mã SAP', width: 120 },
    { key: 'tenVatTu', label: 'Tên vật tư', width: 250 },
    { key: 'dvt', label: 'ĐVT', width: 80 },
    { key: 'khoiLuongTong', label: 'Khối lượng tổng', width: 140 },
    { key: 'thanhTien', label: 'Thành tiền', width: 150 },
    { key: 'donGiaTrungBinh', label: 'Đơn giá trung bình', width: 140 },
    { key: 'donGiaTrungBinh1Ngay', label: 'Đơn giá TB 1 ngày', width: 160 },
    { key: 'phanLoaiVatTu', label: 'Phân loại vật tư', width: 180 }
  ]

  pws['!cols'] = pcols.map(c => ({ wpx: c.width }))

  let pRowIdx = 1

  pws['A1'] = {
    v: 'DANH SÁCH PHÂN NHÓM VẬT TƯ & ĐƠN GIÁ',
    t: 's',
    s: {
      font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0A3D73' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }
  }
  pws['A2'] = {
    v: `Xuất ngày: ${new Date().toLocaleDateString('vi-VN')} | Tổng số dòng: ${(materialPriceRows || []).length}`,
    t: 's',
    s: {
      font: { name: 'Segoe UI', sz: 10, italic: true },
      alignment: { horizontal: 'left', vertical: 'center' }
    }
  }

  pRowIdx = 4

  pcols.forEach((col, colIdx) => {
    const colChar = String.fromCharCode(65 + colIdx)
    const cellRef = `${colChar}${pRowIdx}`
    pws[cellRef] = {
      v: col.label,
      t: 's',
      s: {
        fill: { patternType: 'solid', fgColor: { rgb: '0F58A7' } },
        font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: {
          top: { style: 'thin', color: { rgb: '0A3D73' } },
          bottom: { style: 'medium', color: { rgb: '0A3D73' } },
          left: { style: 'thin', color: { rgb: '0A3D73' } },
          right: { style: 'thin', color: { rgb: '0A3D73' } }
        }
      }
    }
  })

  const rowsToUse = materialPriceRows || []
  rowsToUse.forEach((row, rowIndex) => {
    pRowIdx++
    const isEven = (rowIndex % 2 === 1)
    const rowBgColor = isEven ? 'F8FAFC' : 'FFFFFF'
    
    let meta = { tenVatTu: '—', dvt: '—' }
    if (materialMetadataMap) {
      const sapKey = String(row.maSAP || '').trim().toLowerCase()
      if (materialMetadataMap.get) {
        meta = materialMetadataMap.get(sapKey) || meta
      } else {
        meta = materialMetadataMap[sapKey] || meta
      }
    }

    pcols.forEach((col, colIdx) => {
      const colChar = String.fromCharCode(65 + colIdx)
      const cellRef = `${colChar}${pRowIdx}`

      let val = ''
      let cellType = 's'
      let numFormat = undefined

      if (col.key === 'STT') {
        val = rowIndex + 1
        cellType = 'n'
      } else if (col.key === 'maSAP') {
        val = row.maSAP || ''
        cellType = 's'
      } else if (col.key === 'tenVatTu') {
        val = meta.tenVatTu || ''
        cellType = 's'
      } else if (col.key === 'dvt') {
        val = meta.dvt || ''
        cellType = 's'
      } else if (col.key === 'khoiLuongTong') {
        val = Number(row.khoiLuongTong) || 0
        cellType = 'n'
        numFormat = '#,##0.00'
      } else if (col.key === 'thanhTien') {
        val = Number(row.thanhTien) || 0
        cellType = 'n'
        numFormat = '#,##0'
      } else if (col.key === 'donGiaTrungBinh') {
        val = Number(row.donGiaTrungBinh) || 0
        cellType = 'n'
        numFormat = '#,##0'
      } else if (col.key === 'donGiaTrungBinh1Ngay') {
        val = Number(row.donGiaTrungBinh1Ngay) || 0
        cellType = 'n'
        numFormat = '#,##0'
      } else if (col.key === 'phanLoaiVatTu') {
        val = row.phanLoaiVatTu || 'Chưa phân loại'
        cellType = 's'
      }

      const isCentered = ['STT', 'maSAP', 'dvt', 'phanLoaiVatTu'].includes(col.key)
      const isRight = ['khoiLuongTong', 'thanhTien', 'donGiaTrungBinh', 'donGiaTrungBinh1Ngay'].includes(col.key)

      const cellStyle = {
        font: { name: 'Segoe UI', sz: 9, color: { rgb: '1A1A1A' } },
        alignment: {
          horizontal: isCentered ? 'center' : (isRight ? 'right' : 'left'),
          vertical: 'center',
          wrapText: true
        },
        fill: { patternType: 'solid', fgColor: { rgb: rowBgColor } },
        border: {
          top: { style: 'thin', color: { rgb: 'E2E8F0' } },
          bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
          left: { style: 'thin', color: { rgb: 'E2E8F0' } },
          right: { style: 'thin', color: { rgb: 'E2E8F0' } }
        }
      }

      if (col.key === 'phanLoaiVatTu' && row.phanLoaiVatTu) {
        const text = String(row.phanLoaiVatTu).trim().toLowerCase()
        if (text.includes('tiêu hao')) {
          cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } }
          cellStyle.font.color = { rgb: '047857' }
          cellStyle.font.bold = true
        } else if (text.includes('khấu hao') || text.includes('tài sản')) {
          cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } }
          cellStyle.font.color = { rgb: 'EA580C' }
          cellStyle.font.bold = true
        } else {
          cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } }
          cellStyle.font.color = { rgb: '1D4ED8' }
          cellStyle.font.bold = true
        }
      }

      const cellObj = { v: val, t: cellType, s: cellStyle }
      if (numFormat) cellObj.z = numFormat
      pws[cellRef] = cellObj
    })
  })

  pws['!ref'] = `A1:I${pRowIdx}`
  return pws
}
