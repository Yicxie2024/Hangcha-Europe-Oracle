"""Regression tests for HCE Oracle 1.5.34 (no Windows GUI required)."""
import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from gui_scanner import ScanRow, scan_one
from gui_archive import ArchiveStore, path_key, file_sha256, plan_archive
from gui_archive_repair import plan_historical_delivery_repair
from gui_scan_cache import ScanCache, incremental_scan, changed_since_scan
from gui_tree import matches_filter
from gui import _row_identifier


def row(path, kind, contract='267046', number='', status='识别成功'):
    sha = file_sha256(path)
    return ScanRow(kind, 'Example GmbH', contract, number if kind == '发票' else '',
                   '06.11.2026', '100', '85', status, '示例', str(path), sha,
                   delivery_no=number if kind == '发货单' else '')


class ScanAndArchiveTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.cache = ScanCache(self.root/'cache.db', 'xyc')
        self.store = ArchiveStore(self.root/'archive.db', 'xyc')
        self.counter = {'reads':0}

    def fake_scan(self, path, root, hint, salesperson):
        self.counter['reads'] += 1
        kind = '发货单' if path.name.startswith('delivery') else ('发票' if path.name.startswith('invoice') else '订单')
        n = path.read_text('utf-8')
        return ScanRow(kind, 'A', '267046', 'H20260001' if kind == '发票' else '',
                       '06.11.2026', '100', '85', '识别成功', n,
                       str(path), file_sha256(path),
                       delivery_no='DEL-123' if kind == '发货单' else '')

    def scan(self, cache=None, mode='incremental', store=None, profile='xyc'):
        from gui_scan_cache import incremental_scan
        c = cache or self.cache
        records = (store or self.store).all_archived()
        return incremental_scan([{'path':str(self.root),'hint':'自动识别'}],
                 profile, [], None, None, c, records, mode)

    def test_new_then_cached_then_modified_then_full(self):
        a = self.root/'order.pdf'
        a.write_text('Version A')
        with patch('gui_scan_cache.scan_one',self.fake_scan):
            first, stats, metadata = self.scan()
            self.assertEqual(stats['reparsed'],1)
            self.assertEqual(self.counter['reads'],1)
            self.assertEqual(changed_since_scan(metadata),[])
            second, stats,_ = self.scan()
            self.assertEqual(stats['cached'],1)
            self.assertEqual(self.counter['reads'],1)
            self.assertEqual(second[0].note,'Version A')
            a.write_text('Version B longer')
            self.assertEqual(len(changed_since_scan(metadata)),1)
            third, stats,_ = self.scan()
            self.assertEqual(stats['reparsed'],1)
            self.assertEqual(third[0].note,'Version B longer')
            fourth, stats,_ = self.scan(mode='full')
            self.assertEqual(stats['reparsed'],1)
            self.assertEqual(self.counter['reads'],3)

    def test_profiles_separate(self):
        (self.root/'order.pdf').write_text('A')
        with patch('gui_scan_cache.scan_one',self.fake_scan):
            _, s,_=self.scan()
            self.assertEqual(s['reparsed'],1)
            other=ScanCache(self.root/'cache.db','other-employee')
            _, s,_=self.scan(cache=other)
            self.assertEqual(s['reparsed'],1)
            _, s,_=self.scan()
            self.assertEqual(s['cached'],1)

    def test_archived_pdf_uses_cache_and_new_pdf_is_not_hidden(self):
        o=self.root/'order.pdf'; i=self.root/'invoice.pdf'; d=self.root/'delivery.pdf'
        for f in (o,i,d): f.write_text('data '+f.name)
        with patch('gui_scan_cache.scan_one',self.fake_scan):
            rows, stats,_=self.scan()
            self.assertEqual(stats['reparsed'],3)
            order=next(r for r in rows if r.kind=='订单')
            invoice=next(r for r in rows if r.kind=='发票')
            self.store.archive([('matched',[order,invoice])])
            self.counter['reads']=0
            rows,stats,_=self.scan()
            self.assertEqual(stats['archived_cached'],2)
            self.assertEqual(stats['cached'],1)
            self.assertEqual(self.counter['reads'],0)
            self.assertEqual(len(self.store.archived_for(rows)),2)
            # An archived path replaced with different bytes is NOT hidden.
            o.write_text('replacement content')
            rows,stats,_=self.scan()
            self.assertEqual(stats['archived_changed'],1)
            self.assertNotIn(path_key(o),self.store.archived_for(rows))
            self.assertIn(path_key(i),self.store.archived_for(rows))

    def test_verification_hashes_archived_even_when_metadata_is_constant(self):
        o=self.root/'order.pdf'; i=self.root/'invoice.pdf'
        o.write_text('order'); i.write_text('invoice')
        with patch('gui_scan_cache.scan_one',self.fake_scan):
            rows,stats,_=self.scan()
            self.store.archive([('matched',rows)])
            with patch('gui_scan_cache.file_sha256',wraps=file_sha256) as sha:
                _,stats,_=self.scan(mode='verify')
                self.assertEqual(stats['archive_verified'],2)
                self.assertEqual(sha.call_count,2)

    def test_history_batch_repair_and_restore_as_group(self):
        o=self.root/'order.pdf'; i=self.root/'invoice.pdf'; d=self.root/'delivery.pdf'
        for p in (o,i,d): p.write_text('something '+p.name)
        o_r=row(o,'订单'); i_r=row(i,'发票',number='INV-100'); d_r=row(d,'发货单',number='DN-200')
        self.store.archive([('matched',[o_r,i_r])])
        groups, unresolved = plan_historical_delivery_repair([o_r,i_r,d_r],
                        self.store.all_archived(),{path_key(o),path_key(i)},lambda _:False)
        self.assertEqual(len(groups),1)
        self.assertFalse(unresolved)
        self.assertEqual(groups[0][1][0].delivery_no,'DN-200')
        self.store.archive(groups)
        archive=self.store.all_archived()
        self.assertEqual(len(archive),3)
        self.assertEqual(len({item.group_id for item in archive}),1)
        expanded,bad=self.store.check_restore([path_key(o)])
        self.assertEqual(len(expanded),3)
        self.assertFalse(bad)
        self.store.restore([path_key(o)])
        self.assertFalse(self.store.all_archived())
        self.assertTrue(all(p.exists() for p in (o,i,d)))

    def test_undo_followup_preserves_older_order_and_invoice_archive(self):
        o=self.root/'order.pdf'; i=self.root/'invoice.pdf'; d=self.root/'delivery.pdf'
        for p in (o,i,d): p.write_text('original:'+p.name)
        order=row(o,'订单'); invoice=row(i,'发票',number='INV'); delivery=row(d,'发货单',number='DN')
        self.store.archive([('matched',[order,invoice])])
        added=self.store.archive([('delivery_followup',[delivery])])
        self.assertEqual(len(added),1)
        self.assertEqual(len(self.store.all_archived()),3)
        self.assertEqual(len(self.store.check_undo_last_archive(added)[0]),1)
        self.store.undo_last_archive(added)
        archived=self.store.all_archived()
        self.assertEqual(len(archived),2)
        self.assertEqual({r.kind for r in archived},{'订单','发票'})
        self.assertTrue(d.exists())

    def test_legacy_pending_order_can_receive_shipment_proof_safely(self):
        o=self.root/'order.pdf'; d=self.root/'delivery.pdf'
        o.write_text('order proof'); d.write_text('delivery proof')
        pending=row(o,'订单',status='待开票')
        self.store.archive([('pending',[pending])])
        delivery=row(d,'发货单',number='DN-PEND')
        groups, unresolved=plan_historical_delivery_repair([pending,delivery],
                            self.store.all_archived(),{path_key(o)},lambda _:False)
        self.assertEqual(len(groups),1)
        self.assertFalse(unresolved)
        self.store.archive(groups)
        self.assertEqual(len(self.store.all_archived()),2)
        self.assertEqual(len({item.group_id for item in self.store.all_archived()}),1)

    def test_cache_does_not_freeze_dynamic_pending_invoice_status(self):
        from gui_scanner import mark_pending_invoice_status
        o=self.root/'order.pdf'; o.write_text('order')
        with patch('gui_scan_cache.scan_one',self.fake_scan):
            rows, stats,_=self.scan()
            mark_pending_invoice_status(rows)
            self.assertEqual(rows[0].status,'待开票')
            inv=self.root/'invoice.pdf'; inv.write_text('invoice')
            rows, stats,_=self.scan()
            self.assertEqual(stats['cached'],1)
            self.assertEqual(stats['reparsed'],1)
            mark_pending_invoice_status(rows)
            self.assertEqual(next(r.status for r in rows if r.kind=='订单'),'识别成功')

    def test_ambiguous_or_unreviewed_do_not_auto_archive(self):
        o=self.root/'order.pdf'; i=self.root/'invoice.pdf'; d=self.root/'delivery.pdf'
        for p in (o,i,d): p.write_text(p.name)
        self.store.archive([('matched',[row(o,'订单'),row(i,'发票',number='I')])])
        r=row(d,'发货单',number='DN',status='待复核')
        groups,fail = plan_historical_delivery_repair([r],self.store.all_archived(),set(),lambda _:False)
        self.assertFalse(groups)
        self.assertIn('待复核',fail[0][0])
        r.status='识别成功'; r.contract=''
        groups,fail = plan_historical_delivery_repair([r],self.store.all_archived(),set(),lambda _:False)
        self.assertFalse(groups)
        self.assertIn('订单号未识别',fail[0][0])
        r.contract='267046'
        active_o=row(o,'订单')
        groups,fail = plan_historical_delivery_repair([active_o,r],self.store.all_archived(),set(),lambda _:False)
        self.assertFalse(groups)
        self.assertIn('未归档订单',fail[0][0])

    def test_own_document_number_and_search(self):
        p=self.root/'delivery.pdf'; p.write_text('x')
        r=row(p,'发货单',contract='267046',number='DN-200')
        self.assertEqual(_row_identifier(r),'DN-200')
        self.assertTrue(matches_filter(r,'全部','dn-200'))
        r.delivery_no=''
        self.assertEqual(_row_identifier(r),'—')
        self.assertNotEqual(_row_identifier(r),'267046')


if __name__=='__main__':
    unittest.main(verbosity=2)
