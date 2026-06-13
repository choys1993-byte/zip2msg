#!/usr/bin/env python3
"""
ZIP to MSG Converter
같은 폴더의 .zip 파일을 전부 첨부파일이 포함된 .msg 파일로 변환합니다.
"""

import struct, os, sys, glob
from datetime import datetime, timezone

ENDOFCHAIN  = 0xFFFFFFFE
FREESECT    = 0xFFFFFFFF
NOSTREAM    = 0xFFFFFFFF
FATSECT     = 0xFFFFFFFD
MINI_CUTOFF = 4096


def get_exe_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def encode16(s):
    return (s + '\x00').encode('utf-16-le')


def pad512(data):
    r = len(data) % 512
    return data + b'\x00' * (512 - r) if r else data


def pad_stream(data):
    """4096바이트 배수로 패딩 (minifat 우회)"""
    r = len(data) % MINI_CUTOFF
    return data + b'\x00' * (MINI_CUTOFF - r) if r else data


def filetime_now():
    dt = datetime.now(timezone.utc)
    epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
    return struct.pack('<Q', int((dt - epoch).total_seconds() * 10_000_000))


def de(name, etype, start, size,
       child=NOSTREAM, left=NOSTREAM, right=NOSTREAM, color=1):
    """OLE2 디렉토리 엔트리 128바이트"""
    ne   = name.encode('utf-16-le')[:62]
    nlen = len(ne) + 2 if ne else 0
    ne   = ne.ljust(64, b'\x00')
    b    = struct.pack('<H', nlen) + struct.pack('<B', etype) + struct.pack('<B', color)
    b   += struct.pack('<III', left, right, child)
    b   += b'\x00' * 16 + struct.pack('<IQQ', 0, 0, 0)
    b   += struct.pack('<II', start if start != NOSTREAM else ENDOFCHAIN, size)
    b   += b'\x00' * 4
    assert len(ne) + len(b) == 128
    return ne + b


def make_props(*entries):
    """
    프로퍼티 스트림 생성
    entries: (tag, ptype, value)
      - fixed: value = bytes (8바이트)
      - variable: value = int (스트림 크기)
    """
    out = b'\x00' * 8  # 헤더 reserved
    for tag, ptype, val in entries:
        if isinstance(val, bytes):
            out += struct.pack('<HH', tag, ptype) + val + b'\x00' * 4
        else:
            out += struct.pack('<HHI', tag, ptype, val) + b'\x00' * 4
    return pad_stream(out)


def build_zip_msg(zip_path):
    """zip 파일을 첨부파일로 포함한 MSG 바이너리 생성"""
    zip_name = os.path.basename(zip_path)
    zip_ext  = os.path.splitext(zip_name)[1]

    with open(zip_path, 'rb') as f:
        zip_raw = f.read()
    zip_padded = pad_stream(zip_raw)

    # ── 스트림 데이터 준비 ──
    mc_p    = pad_stream(encode16('IPM.Note'))
    subj_p  = pad_stream(encode16(zip_name))
    ext_p   = pad_stream(encode16(zip_ext))
    lname_p = pad_stream(encode16(zip_name))
    dname_p = pad_stream(encode16(zip_name))

    root_prop = make_props(
        (0x0039, 0x0040, filetime_now()),            # PR_CLIENT_SUBMIT_TIME
        (0x001A, 0x001F, len(mc_p)),                 # PR_MESSAGE_CLASS
        (0x0037, 0x001F, len(subj_p)),               # PR_SUBJECT
        (0x0E1B, 0x000B, struct.pack('<II', 1, 0)),  # PR_HASATTACH
    )
    att_prop = make_props(
        (0x3705, 0x0003, struct.pack('<II', 1, 0)),  # PR_ATTACH_METHOD = by value
        (0x0E21, 0x0003, struct.pack('<II', 0, 0)),  # PR_ATTACH_NUM
        (0x3704, 0x001F, len(ext_p)),                # PR_ATTACH_EXTENSION
        (0x3707, 0x001F, len(lname_p)),              # PR_ATTACH_LONG_FILENAME
        (0x3001, 0x001F, len(dname_p)),              # PR_DISPLAY_NAME
        (0x3701, 0x0102, len(zip_padded)),           # PR_ATTACH_DATA_BIN
    )

    # ── 섹터 할당 ──
    sectors = []
    fat     = []

    def alloc(data):
        data  = pad512(data)
        nsec  = len(data) // 512
        start = len(sectors)
        for i in range(nsec):
            sectors.append(data[i * 512:(i + 1) * 512])
            fat.append(start + i + 1 if i < nsec - 1 else ENDOFCHAIN)
        return start, len(data)

    rp_s, rp_sz = alloc(root_prop)
    mc_s, mc_sz = alloc(mc_p)
    sb_s, sb_sz = alloc(subj_p)
    ap_s, ap_sz = alloc(att_prop)
    ex_s, ex_sz = alloc(ext_p)
    ln_s, ln_sz = alloc(lname_p)
    dn_s, dn_sz = alloc(dname_p)
    at_s, at_sz = alloc(zip_padded)

    # ── 디렉토리 엔트리 ──
    # SID 0: Root Entry                child=1
    # SID 1: __properties (root)       right=2
    # SID 2: __substg MC               right=3
    # SID 3: __substg Subject          right=4
    # SID 4: __attach_#00000000        child=5  (storage)
    # SID 5: __properties (attach)     right=6
    # SID 6: __substg extension        right=7
    # SID 7: __substg long filename    right=8
    # SID 8: __substg display name     right=9
    # SID 9: __substg attach data

    dir_b  = de('Root Entry',                   5, NOSTREAM, 0,     child=1, color=0)
    dir_b += de('__properties_version1.0',      2, rp_s, rp_sz,    right=2)
    dir_b += de('__substg1.0_001A001F',          2, mc_s, mc_sz,    right=3)
    dir_b += de('__substg1.0_0037001F',          2, sb_s, sb_sz,    right=4)
    dir_b += de('__attach_version1.0_#00000000', 1, NOSTREAM, 0,    child=5)
    dir_b += de('__properties_version1.0',      2, ap_s, ap_sz,    right=6)
    dir_b += de('__substg1.0_3704001F',          2, ex_s, ex_sz,    right=7)
    dir_b += de('__substg1.0_3707001F',          2, ln_s, ln_sz,    right=8)
    dir_b += de('__substg1.0_3001001F',          2, dn_s, dn_sz,    right=9)
    dir_b += de('__substg1.0_37010102',          2, at_s, at_sz)

    dir_s, _ = alloc(pad512(dir_b))

    # ── FAT 섹터 ──
    fat_idx  = len(sectors)
    rem      = 128 - (len(fat) % 128)
    fat_full = fat + ([FREESECT] * rem if rem < 128 else [])
    if len(fat_full) <= fat_idx:
        fat_full += [FREESECT] * (fat_idx - len(fat_full) + 1)
    fat_full[fat_idx] = FATSECT
    sectors.append(struct.pack(f'<{min(128, len(fat_full))}I', *fat_full[:128]))

    # ── OLE2 헤더 ──
    hdr = bytearray(512)
    hdr[0:8] = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
    struct.pack_into('<H', hdr, 24, 0x3E)
    struct.pack_into('<H', hdr, 26, 3)
    struct.pack_into('<H', hdr, 28, 0xFFFE)
    struct.pack_into('<H', hdr, 30, 9)
    struct.pack_into('<H', hdr, 32, 6)
    struct.pack_into('<I', hdr, 40, 0)
    struct.pack_into('<I', hdr, 44, 1)
    struct.pack_into('<I', hdr, 48, dir_s)
    struct.pack_into('<I', hdr, 52, 0)
    struct.pack_into('<I', hdr, 56, 0x1000)
    struct.pack_into('<I', hdr, 60, ENDOFCHAIN)
    struct.pack_into('<I', hdr, 64, 0)
    struct.pack_into('<I', hdr, 68, ENDOFCHAIN)
    struct.pack_into('<I', hdr, 72, 0)
    struct.pack_into('<I', hdr, 76, fat_idx)
    for i in range(1, 109):
        struct.pack_into('<I', hdr, 76 + i * 4, FREESECT)

    return bytes(hdr) + b''.join(sectors)


def main():
    base_dir  = get_exe_dir()
    zip_files = glob.glob(os.path.join(base_dir, '*.zip'))

    print('=' * 52)
    print('  ZIP → MSG Converter')
    print('=' * 52)

    if not zip_files:
        print(f'\n[오류] 같은 폴더에 .zip 파일이 없습니다.')
        print(f'  폴더: {base_dir}')
        input('\nEnter 키를 눌러 종료...')
        sys.exit(1)

    print(f'\n  {len(zip_files)}개 파일 발견\n')

    ok = fail = 0
    for zip_path in zip_files:
        base     = os.path.splitext(zip_path)[0]
        msg_path = base + '.msg'
        fname    = os.path.basename(zip_path)
        try:
            out = build_zip_msg(zip_path)
            with open(msg_path, 'wb') as f:
                f.write(out)
            print(f'  ✓  {fname}  →  {os.path.basename(msg_path)}')
            ok += 1
        except Exception as e:
            print(f'  ✗  {fname}  ({e})')
            fail += 1

    print(f'\n  완료: 성공 {ok}개 / 실패 {fail}개')
    input('\nEnter 키를 눌러 종료...')


if __name__ == '__main__':
    main()
