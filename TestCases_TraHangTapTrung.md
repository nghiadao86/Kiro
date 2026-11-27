# Test Cases – Giải pháp trả hàng tập trung (AX 2012 R3)

**Phạm vi:** Trả hàng **cùng store** xuất bill (Store B = Store A) và **khác store** xuất bill (Store B ≠ Store A).
**Nguồn:** `Giải pháp trả hàng tập trung.docx` (API Return Items, End quarantine, Get ticket info, Create payment journal, Batch job trả hàng).

**Quy ước**
- Store A = store xuất bill (OriginStore). Store B = store nhận trả hàng (Store trong input).
- VTB = kho tập trung. RA = team kiểm hàng. SO = Sales Order. TF = Transfer Order. PTTT = phương thức thanh toán.
- Ưu tiên: P1 = bắt buộc, P2 = quan trọng, P3 = bổ sung.
- Cột "Kết quả mong đợi" kiểm tra cả dữ liệu staging, chứng từ AX, tồn kho và GL.

## 0. Điều kiện tiên quyết / dữ liệu chung
| Mã | Nội dung |
|---|---|
| PRE-01 | Có Store A, Store B (khác nhau), kho tương ứng, kho VTB. Cấu hình quarantine cho các kho. |
| PRE-02 | Khách hàng KH-000001; bill offline đã có Retail transaction/sales trans (có `recId` dùng làm LineRecId). |
| PRE-03 | Item có quản lý batch. Có bill gồm ≥2 line; line có discount. |
| PRE-04 | Form *Payment method of store* có cấu hình Account cho Tiền mặt, Chuyển khoản tại Store A và B. |
| PRE-05 | Form *Methods of payment (AR)* có cấu hình Account cho Coupon, P-xu, Phí trả hàng (ledger 113500). |
| PRE-06 | Batch job xử lý trả hàng đã cấu hình. User/role cho Web App, RA, Kế toán AR. |
| PRE-07 | Có bill offline có SalesId (Sales order có giá trị) và bill offline không có SalesId (null). |

## A. Validate API Return Items (input)
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-A01 | Gọi API đủ trường bắt buộc | Gửi input hợp lệ (TicketId, OrderId, Source, TransDate, StoreId, Lines, Payments) | API trả thành công; tạo bản ghi StagingTable/Line/Payment theo TicketId | P1 |
| TC-A02 | Thiếu TicketId | Bỏ TicketId | API trả lỗi rõ ràng; không ghi staging | P1 |
| TC-A03 | Thiếu từng trường Mandatory khác | Lần lượt bỏ OrderId, Source, TransDate, StoreId, ItemId, Unit, LineRecId, Qty, TenderType, PaymentAmount, Currency | Mỗi lần API từ chối, thông báo đúng trường thiếu | P1 |
| TC-A04 | TicketId trùng (idempotency) | Gọi lại API với cùng TicketId | Không tạo trùng SO/TF/quarantine/journal; trả kết quả/lỗi xác định | P1 |
| TC-A05 | Vượt độ dài trường | TicketId >50, OrderId >20, StoreId >10, Description >255, Staff >20, Unit >10, ItemId >20 | Bị từ chối hoặc báo lỗi, không cắt dữ liệu âm thầm | P2 |
| TC-A06 | Source không hợp lệ | Source ngoài Offline/Online | Lỗi validate | P2 |
| TC-A07 | Source = Online | Gửi Source=Online | Xử lý theo đúng quy tắc (tài liệu chỉ mô tả Offline → xác nhận với BA) | P2 |
| TC-A08 | Qty dương | Qty > 0 (tài liệu: SL trả là âm) | Lỗi/chuẩn hóa theo thiết kế; không tạo chứng từ sai dấu | P1 |
| TC-A09 | Qty vượt SL đã bán | Qty trả > SL trên sales trans LineRecId | Bị từ chối | P1 |
| TC-A10 | Trả nhiều lần cùng line | Trả 2 lần, tổng ≤ SL bán; lần 3 vượt | Lần 1–2 OK; lần 3 bị từ chối do vượt số lượng còn lại | P1 |
| TC-A11 | LineRecId không tồn tại / không thuộc OrderId | LineRecId sai | Lỗi, không ghi chứng từ | P1 |
| TC-A12 | StoreId không tồn tại | StoreId sai | Lỗi | P2 |
| TC-A13 | Currency ≠ VND | Currency = USD | Bị từ chối hoặc xử lý theo thiết kế | P3 |
| TC-A14 | TenderType không hợp lệ | TenderType chưa cấu hình | Lỗi, không tạo journal | P2 |
| TC-A15 | Tổng PaymentAmount ≠ tổng tiền trả | Lệch số tiền | Lỗi/cảnh báo theo thiết kế | P1 |
| TC-A16 | TransDate kỳ kế toán đóng | TransDate thuộc kỳ GL/AR đã đóng | Lỗi rõ ràng, không post | P2 |
| TC-A17 | Description/Staff để trống | Bỏ qua 2 trường tùy chọn | API thành công | P3 |
| TC-A18 | Qty = 0 | Qty=0 | Bị từ chối | P3 |

## B. Trả hàng CÙNG store (Store B = Store A)
### B1. Sales order có giá trị → Credit note
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-B01 | Trả toàn bộ 1 line | Gọi API, Source=Offline, SO có giá trị | Tạo credit note trên SO, line từ sales trans recId=LineRecId, SalesQty = -Qty | P1 |
| TC-B02 | Trả một phần SL | Qty < SL bán | Credit note đúng SL; giá/discount tính theo tỷ lệ | P1 |
| TC-B03 | Trả nhiều line | 2–3 line | Mỗi line một dòng credit note đúng ItemId, Unit, Qty, Batch | P1 |
| TC-B04 | Packing slip credit note | Sau khi tạo credit note | Packed; tạo quarantine tại kho Store A | P1 |
| TC-B05 | Quarantine auto "Reversed" | Kiểm tra quarantine order | Quarantine tạo ra auto check Reversed (cần xác nhận "Reversed"/"Reserved") | P1 |
| TC-B06 | Cập nhật staging | Kiểm tra StagingTable/Line | SalesId, QuarantineOrder được cập nhật | P1 |
| TC-B07 | Tồn kho | Kiểm tra on-hand sau trả | Hàng nằm trong quarantine, chưa vào tồn khả dụng; đúng batch | P1 |
| TC-B08 | Không tạo TF | Kiểm tra Transfer order | Không có TF cho trả cùng store | P1 |
| TC-B09 | Post invoice credit note | Sau packing | Invoice posted; InvoiceId ghi vào StagingTable; doanh thu/thuế giảm đúng | P1 |

### B2. Sales order rỗng/null → SO âm
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-B10 | Trả hàng bill không có SalesId | Source=Offline, SO null | Tạo SO âm với line từ sales trans recId=LineRecId, SalesQty=-Qty | P1 |
| TC-B11 | Thông tin SO âm | Kiểm tra header/line | CustAccount, item, qty, discountAmount, netAmount khớp input/bill | P1 |
| TC-B12 | SO âm packing + quarantine + post | Kiểm tra | Packing slip, quarantine (Reversed), invoice posted, SalesId/InvoiceId cập nhật | P1 |

### B3. Thanh toán / hoàn tiền (cùng store)
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-B13 | Chỉ tiền mặt | Payments: Tiền mặt 100k | AR payment journal: Dr Cust 100k / Cr Bank (account của payment method của store trả); post; settlement với SO; JournalId, Voucher lưu StagingPayment | P1 |
| TC-B14 | Tiền mặt + Phí trả hàng | Tiền mặt 80k + Phí trả hàng 20k (bill 100k) | Journal: Dr Cust 100k / Cr Bank 80k / Cr Ledger 113500 20k | P1 |
| TC-B15 | Tiền mặt + coupon/P-xu | Có coupon/P-xu | Line coupon/P-xu lấy account từ form Methods of payment (AR) theo payment method & store | P1 |
| TC-B16 | Chỉ chuyển khoản | Payments chỉ CK | Không tạo payment journal lúc gọi API Return Items; StagingPayment lưu, JournalId trống | P1 |
| TC-B17 | Hỗn hợp CK + khác | CK 30k + Tiền mặt 50k + Phí 20k | Journal tự tạo cho phần khác CK: Dr Cust 70k / Cr Bank 50k / Cr Ledger 20k; phần CK 30k chờ API createPaymentJournal | P1 |
| TC-B18 | Settlement | Kiểm tra open transaction khách | Journal settle với SO/credit note, công nợ cấn trừ đúng | P1 |
| TC-B19 | Bút toán chuyển tiền store | Có tiền mặt, Store B=A | Kiểm tra theo thiết kế (cùng store: không cần chuyển tiền giữa store – xác nhận BA) | P2 |
| TC-B20 | TransDate journal | Kiểm tra TransDate | = TransDate input | P1 |
| TC-B21 | Làm tròn tiền | Số lẻ | Không lệch debit/credit; journal cân | P2 |

## C. Trả hàng KHÁC store (Store B ≠ Store A)
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-C01 | Credit note (SO có giá trị) khác store | Trả tại B hàng bán từ A | Tạo credit note, packing slip, quarantine tại kho A | P1 |
| TC-C02 | SO âm (SO null) khác store | Như trên với bill không SalesId | Tạo SO âm, packing, quarantine tại kho A | P1 |
| TC-C03 | End quarantine kho A | Sau bước C01/C02 | Quarantine kho A được End | P1 |
| TC-C04 | Tạo TF A → B | Kiểm tra TF | TF từ kho A sang kho B đúng item, batch, qty; kho A giảm, kho B tăng | P1 |
| TC-C05 | Quarantine kho B | Kiểm tra | Quarantine tạo ở kho B, auto check Reserved; chờ RA check | P1 |
| TC-C06 | Cập nhật staging | Kiểm tra | SalesId, TransferOrder, QuarantineOrder (kho B) được lưu | P1 |
| TC-C07 | Post invoice | Kiểm tra | InvoiceId lưu staging; doanh thu ghi nhận đúng store/dimension của bill gốc (A) | P1 |
| TC-C08 | Tồn kho tổng thể | So sánh on-hand A, B trước/sau | A giảm, B tăng đúng SL; tổng hệ thống không lệch | P1 |
| TC-C09 | Hoàn tiền tiền mặt | Payment tiền mặt tại B | Journal Dr Cust / Cr Bank (account theo payment method + store B); post; settle | P1 |
| TC-C10 | Bút toán chuyển tiền B → A | Có tiền mặt | Journal: Dr Bank STO-A 50k / Cr Bank STO-B 50k; đã post | P1 |
| TC-C11 | Không có tiền mặt | Chỉ CK / chỉ coupon | Không tạo bút toán chuyển tiền store | P1 |
| TC-C12 | Số tiền chuyển store | Tiền mặt + CK + phí | Số chuyển = đúng phần tiền mặt thực trả | P1 |
| TC-C13 | Hỗn hợp có CK | CK 30k + khác | Journal phần khác CK tạo ngay; CK chờ API | P1 |
| TC-C14 | Trả nhiều line, nhiều batch | 2+ line, batch khác nhau | TF và quarantine giữ đúng batch từng line | P1 |
| TC-C15 | Lỗi khi tạo TF | Kho A không đủ/không tồn tại hoặc cấu hình thiếu | Lỗi được ghi, không để dữ liệu nửa chừng; có thể xử lý lại | P1 |
| TC-C16 | Rollback/Retry | Giả lập lỗi sau bước End quarantine | Retry không tạo trùng chứng từ | P1 |

## D. Create payment journal (PTTT chuyển khoản)
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-D01 | Tạo journal CK | Gọi API createPaymentJournal (TicketId, TransDate) | Dr Cust 30k / Cr Ledger (account payment method CK của store trả) 30k; post | P1 |
| TC-D02 | Settlement | Kiểm tra | Mark settlement với SO trong StagingTable | P1 |
| TC-D03 | Cập nhật staging | Kiểm tra | JournalId, Voucher cập nhật StagingPayment | P1 |
| TC-D04 | TransDate | Truyền TransDate khác ngày trả | Journal dùng đúng TransDate input | P1 |
| TC-D05 | Gọi lặp | Gọi 2 lần cùng TicketId | Lần 2 không tạo journal trùng | P1 |
| TC-D06 | TicketId không tồn tại | TicketId sai | Lỗi | P2 |
| TC-D07 | Ticket không có CK | Ticket chỉ tiền mặt | Không tạo journal; thông báo phù hợp | P2 |
| TC-D08 | Gọi trước khi SO được invoice | Chưa post invoice | Lỗi/chặn theo thiết kế | P2 |
| TC-D09 | Số tiền | Nhiều dòng CK | Số tiền = sum(PaymentAmount) các dòng CK | P1 |
| TC-D10 | Kỳ kế toán đóng | TransDate thuộc kỳ đóng | Lỗi rõ ràng | P2 |

## E. API Get ticket info
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-E01 | Lấy thông tin ticket | Gọi với TicketId hợp lệ | Trả đầy đủ dữ liệu 3 bảng StagingTable, StagingLine, StagingPayment | P1 |
| TC-E02 | Sau từng bước xử lý | Gọi sau khi có SO/TF/Quarantine/Invoice/Journal | Các trường SalesId, InvoiceId, TransferOrder, QuarantineOrder, JournalId, Voucher phản ánh trạng thái đúng | P1 |
| TC-E03 | TicketId không tồn tại | Gọi sai | Lỗi/rỗng, có thông báo | P2 |

## F. API End quarantine
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-F01 | End quarantine hợp lệ | Gọi với QuarantineOrder kho B | Quarantine kết thúc, hàng vào tồn khả dụng | P1 |
| TC-F02 | QuarantineOrder không tồn tại | Gọi sai | Lỗi | P2 |
| TC-F03 | Đã End trước đó | Gọi lại | Không thay đổi, báo lỗi/đã xử lý | P2 |
| TC-F04 | Quyền | User không có quyền RA | Từ chối | P2 |
| TC-F05 | Tồn kho sau End | Kiểm tra on-hand | Tăng đúng SL/batch | P1 |

## G. Batch job xử lý trả hàng
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-G01 | Chạy job cùng store | Có ticket pending, chạy job | SO âm → packing slip + invoice (InvoiceDate = input); hoàn tiền; quarantine Reversed | P1 |
| TC-G02 | Chạy job khác store | Ticket khác store | End quarantine A → TF A→B → quarantine B (Reserved) | P1 |
| TC-G03 | Journal đa dòng | Tiền mặt 50k + CK 30k + coupon/phí 20k | Dr Cust 100k / Cr Bank STO-A 50k / Cr Ledger 113937 30k / Cr Ledger 113500 20k | P1 |
| TC-G04 | Account theo cấu hình | Kiểm tra account | Lấy đúng từ Payment method of store và Methods of payment (AR) theo store trả | P1 |
| TC-G05 | Auto settlement + post | Kiểm tra | Cấn trừ công nợ với SO âm; payment posted | P1 |
| TC-G06 | Chuyển tiền STO-B → STO-A | Có tiền mặt | Dr Bank STO-A / Cr Bank STO-B đúng số tiền | P1 |
| TC-G07 | Job chạy lại | Chạy 2 lần | Không xử lý trùng ticket đã xử lý | P1 |
| TC-G08 | Job lỗi 1 ticket | 1 ticket lỗi trong lô | Các ticket khác vẫn xử lý; lỗi được log để xử lý lại | P1 |
| TC-G09 | Khối lượng lớn | ≥ 500 ticket | Không timeout; kết quả đúng | P3 |
| TC-G10 | Hạch toán doanh thu | Kiểm tra GL | Giảm trừ doanh thu, thuế, giá vốn đúng | P1 |

## H. RA kiểm hàng & xử lý sau quarantine
| ID | Tiêu đề | Bước | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|---|
| TC-H01 | RA End quarantine – hàng đạt | RA kiểm, End quarantine | Hàng vào tồn khả dụng kho B (hoặc A nếu cùng store) | P1 |
| TC-H02 | RA tạo TF B → VTB | Hàng cần về kho tập trung | TF giảm kho B, tăng VTB đúng SL, batch | P1 |
| TC-H03 | RA không chuyển VTB | Giữ tại kho store | Không tạo TF; tồn kho đúng | P2 |
| TC-H04 | Hàng không đạt | Hàng lỗi/hết hạn | Xử lý theo quy trình (xác nhận với BA) | P2 |
| TC-H05 | Phân quyền RA | User không phải RA | Không thực hiện được End quarantine/TF | P2 |

## I. Đổi hàng (tham khảo, ngoài phạm vi chính – chỉ kiểm tra liên quan)
| ID | Tiêu đề | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|
| TC-I01 | Trả hàng không bị ảnh hưởng bởi API Exchange Items | Hai luồng độc lập, không tạo chứng từ chéo | P3 |

## J. Kiểm tra tích hợp / End-to-End
| ID | Kịch bản | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|
| TC-J01 | E2E cùng store, tiền mặt, SO có giá trị | Web → API → credit note → quarantine → invoice → journal → RA End quarantine; số liệu khớp | P1 |
| TC-J02 | E2E cùng store, SO null | Như trên với SO âm | P1 |
| TC-J03 | E2E khác store, tiền mặt | Có TF A→B, chuyển tiền B→A, quarantine B | P1 |
| TC-J04 | E2E khác store, hỗn hợp CK + tiền mặt + phí | Journal tách hai phần; CK qua API createPaymentJournal | P1 |
| TC-J05 | E2E khác store, trả về VTB | Thêm TF B→VTB do RA | P1 |
| TC-J06 | Tra cứu bill → tạo ticket → xác nhận | Web lấy bill qua API (ReceiptId, StoreNumber); gọi API trả hàng sau xác nhận | P1 |
| TC-J07 | Đối chiếu tổng | Tồn kho, công nợ khách, quỹ từng store, GL cân sau toàn bộ luồng | P1 |
| TC-J08 | Mất kết nối giữa chừng | Web gọi API timeout rồi gọi lại | Không trùng chứng từ; trạng thái nhất quán | P1 |

## K. Phi chức năng & bảo mật
| ID | Nội dung | Kết quả mong đợi | Ưu tiên |
|---|---|---|---|
| TC-K01 | Xác thực API | Gọi không token/sai quyền bị từ chối | P1 |
| TC-K02 | Audit log | Có log người/tiện ích gọi, thời gian, TicketId | P2 |
| TC-K03 | Hiệu năng API | Thời gian phản hồi trong ngưỡng đã thỏa thuận | P3 |
| TC-K04 | Xử lý đồng thời | Hai ticket cùng item/batch đồng thời | Không âm tồn, không deadlock | P2 |

## Điểm cần làm rõ với BA / Dev
1. Tài liệu viết cả "Reversed" lẫn "Reserved" cho quarantine; cần xác nhận trạng thái đúng ở cùng store và khác store.
2. Mục Payment của API Return Items: "PTTT trả hàng (Tiền mặt/CK/Phí trả hàng)"; coupon/P-xu có là TenderType riêng không?
3. Trả hàng cùng store có cần bút toán chuyển tiền giữa store không (tài liệu chỉ nêu cho trường hợp có tiền mặt, ví dụ STO-B → STO-A)?
4. Account ledger khác nhau giữa ví dụ (113500 cho phí/coupon, 113937 cho CK trong batch job) – xác nhận quy tắc.
5. Source=Online chưa có mô tả luồng xử lý.
6. Quy tắc validate Qty vượt số lượng còn được trả, hạn trả hàng, kỳ đóng.
7. Mô tả output của các API (Get Bill Info, Get master data, Return Items) còn trống trong tài liệu.
8. Cơ chế idempotency theo TicketId và xử lý khi lỗi giữa chừng (retry/rollback).
9. Hàng khác store: batch job và API trực tiếp – khác biệt thời điểm xử lý? "Trả hàng cùng kho/khác kho" trong batch job có đồng nghĩa với cùng/khác store?
