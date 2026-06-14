#!/usr/bin/env python3
"""
ZIP to MSG Converter
같은 폴더의 .zip 파일을 전부 첨부파일이 포함된 .msg 파일로 변환합니다.
제목, 발신자, 본문 등이 zip 파일명 기반으로 자동 채워집니다.
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


def encode16(s): return (s+'\x00').encode('utf-16-le')

def pad512(d):
    r=len(d)%512; return d+b'\x00'*(512-r) if r else d

def pad_stream(d):
    r=len(d)%MINI_CUTOFF; return d+b'\x00'*(MINI_CUTOFF-r) if r else d

def filetime_now():
    dt=datetime.now(timezone.utc); epoch=datetime(1601,1,1,tzinfo=timezone.utc)
    return struct.pack('<Q',int((dt-epoch).total_seconds()*10_000_000))

def de(name, etype, start, size,
       child=NOSTREAM, left=NOSTREAM, right=NOSTREAM, color=1):
    ne=name.encode('utf-16-le')[:62]; nlen=len(ne)+2 if ne else 0
    ne=ne.ljust(64,b'\x00')
    b=struct.pack('<H',nlen)+struct.pack('<B',etype)+struct.pack('<B',color)
    b+=struct.pack('<III',left,right,child)
    b+=b'\x00'*16+struct.pack('<IQQ',0,0,0)
    b+=struct.pack('<II',start if start!=NOSTREAM else ENDOFCHAIN,size)
    b+=b'\x00'*4
    assert len(ne)+len(b)==128
    return ne+b

def empty_de():
    """빈 디렉토리 엔트리 (패딩용)"""
    return b'\x00'*64+struct.pack('<HBBI',0,0,1,NOSTREAM)+\
           struct.pack('<III',NOSTREAM,NOSTREAM,NOSTREAM)+\
           b'\x00'*16+struct.pack('<IQQ',0,0,0)+\
           struct.pack('<II',ENDOFCHAIN,0)+b'\x00'*4

def build_root_props(entries):
    body=b''
    for ptype,tag,val in entries:
        body+=struct.pack('<HH',ptype,tag)+val[:4]
    return pad_stream(b'\x00'*32+body)

def build_att_props(entries):
    body=b''
    for ptype,tag,val in entries:
        body+=struct.pack('<HH',ptype,tag)+val[:4]
    return pad_stream(b'\x00'*8+body)


def build_zip_msg(zip_path):
    zip_name  = os.path.basename(zip_path)
    zip_stem  = os.path.splitext(zip_name)[0]
    zip_ext   = os.path.splitext(zip_name)[1]
    mime_type = 'application/x-zip-compressed'

    # zip 파일명 기반 자동 생성
    sender_name  = zip_stem
    sender_email = f'{zip_stem}@attachment.msg'
    conv_topic   = zip_stem
    body_text    = f'Attachment: {zip_name}'

    with open(zip_path,'rb') as f: zip_raw=f.read()
    zip_padded=pad_stream(zip_raw)

    # 스트림 데이터
    mc_p     = pad_stream(encode16('IPM.Note'))
    subj_p   = pad_stream(encode16(zip_name))
    sname_p  = pad_stream(encode16(sender_name))
    semail_p = pad_stream(encode16(sender_email))
    stype_p  = pad_stream(encode16('SMTP'))
    ctopic_p = pad_stream(encode16(conv_topic))
    body_p   = pad_stream(encode16(body_text))
    dispto_p = pad_stream(encode16(''))
    ext_p    = pad_stream(encode16(zip_ext))
    short_p  = pad_stream(encode16(zip_name))
    lname_p  = pad_stream(encode16(zip_name))
    dname_p  = pad_stream(encode16(zip_name))
    mime_p   = pad_stream(encode16(mime_type))
    locale_p = pad_stream(encode16('EnUs'))

    root_prop=build_root_props([
        (0x0040,0x0039,filetime_now()),
        (0x001F,0x001A,struct.pack('<I',len(mc_p))),
        (0x001F,0x0037,struct.pack('<I',len(subj_p))),
        (0x001F,0x0070,struct.pack('<I',len(ctopic_p))),
        (0x001F,0x0C1A,struct.pack('<I',len(sname_p))),
        (0x001F,0x0C1F,struct.pack('<I',len(semail_p))),
        (0x001F,0x0C1E,struct.pack('<I',len(stype_p))),
        (0x001F,0x0042,struct.pack('<I',len(sname_p))),
        (0x001F,0x0065,struct.pack('<I',len(semail_p))),
        (0x001F,0x0064,struct.pack('<I',len(stype_p))),
        (0x001F,0x0E04,struct.pack('<I',len(dispto_p))),
        (0x001F,0x1000,struct.pack('<I',len(body_p))),
        (0x000B,0x0E1B,struct.pack('<I',1)),
    ])
    att_prop=build_att_props([
        (0x0003,0x0E21,struct.pack('<I',0)),
        (0x0003,0x3705,struct.pack('<I',1)),
        (0x001F,0x3001,struct.pack('<I',len(dname_p))),
        (0x001F,0x3703,struct.pack('<I',len(short_p))),
        (0x001F,0x3704,struct.pack('<I',len(ext_p))),
        (0x001F,0x3707,struct.pack('<I',len(lname_p))),
        (0x001F,0x370E,struct.pack('<I',len(mime_p))),
        (0x001F,0x3A0C,struct.pack('<I',len(locale_p))),
        (0x0102,0x3701,struct.pack('<I',len(zip_padded))),
    ])

    sectors=[]; fat=[]
    def alloc(data):
        data=pad512(data); nsec=len(data)//512; start=len(sectors)
        for i in range(nsec):
            sectors.append(data[i*512:(i+1)*512])
            fat.append(start+i+1 if i<nsec-1 else ENDOFCHAIN)
        return start,len(data)

    rp_s,rp_sz   = alloc(root_prop)
    mc_s,mc_sz   = alloc(mc_p)
    sb_s,sb_sz   = alloc(subj_p)
    sn_s,sn_sz   = alloc(sname_p)
    se_s,se_sz   = alloc(semail_p)
    st_s,st_sz   = alloc(stype_p)
    ct_s,ct_sz   = alloc(ctopic_p)
    bd_s,bd_sz   = alloc(body_p)
    dt_s,dt_sz   = alloc(dispto_p)
    ap_s,ap_sz   = alloc(att_prop)
    ex_s,ex_sz   = alloc(ext_p)
    sh_s,sh_sz   = alloc(short_p)
    ln_s,ln_sz   = alloc(lname_p)
    dn_s,dn_sz   = alloc(dname_p)
    mm_s,mm_sz   = alloc(mime_p)
    lc_s,lc_sz   = alloc(locale_p)
    at_s,at_sz   = alloc(zip_padded)
    ng_s,ng_sz   = alloc(pad_stream(b'\x00'*16))
    ne_s,ne_sz   = alloc(pad_stream(b'\x00'*8))
    ns_s,ns_sz   = alloc(pad_stream(b'\x00'*4))

    # 디렉토리 엔트리
    dir_entries=[]
    dir_entries.append(de('Root Entry',5,NOSTREAM,0,child=1,color=0))
    dir_entries.append(de('__properties_version1.0',2,rp_s,rp_sz,right=2))
    dir_entries.append(de('__substg1.0_001A001F',2,mc_s,mc_sz,right=3))
    dir_entries.append(de('__substg1.0_0037001F',2,sb_s,sb_sz,right=4))
    dir_entries.append(de('__substg1.0_0C1A001F',2,sn_s,sn_sz,right=5))
    dir_entries.append(de('__substg1.0_0C1F001F',2,se_s,se_sz,right=6))
    dir_entries.append(de('__substg1.0_0C1E001F',2,st_s,st_sz,right=7))
    dir_entries.append(de('__substg1.0_0070001F',2,ct_s,ct_sz,right=8))
    dir_entries.append(de('__substg1.0_1000001F',2,bd_s,bd_sz,right=9))
    dir_entries.append(de('__substg1.0_0E04001F',2,dt_s,dt_sz,right=10))
    dir_entries.append(de('__attach_version1.0_#00000000',1,NOSTREAM,0,child=11,right=19))
    dir_entries.append(de('__properties_version1.0',2,ap_s,ap_sz,right=12))
    dir_entries.append(de('__substg1.0_3704001F',2,ex_s,ex_sz,right=13))
    dir_entries.append(de('__substg1.0_3703001F',2,sh_s,sh_sz,right=14))
    dir_entries.append(de('__substg1.0_3707001F',2,ln_s,ln_sz,right=15))
    dir_entries.append(de('__substg1.0_3001001F',2,dn_s,dn_sz,right=16))
    dir_entries.append(de('__substg1.0_370E001F',2,mm_s,mm_sz,right=17))
    dir_entries.append(de('__substg1.0_3A0C001F',2,lc_s,lc_sz,right=18))
    dir_entries.append(de('__substg1.0_37010102',2,at_s,at_sz))
    dir_entries.append(de('__nameid_version1.0',1,NOSTREAM,0,child=20))
    dir_entries.append(de('__substg1.0_00020102',2,ng_s,ng_sz,right=21))
    dir_entries.append(de('__substg1.0_00030102',2,ne_s,ne_sz,right=22))
    dir_entries.append(de('__substg1.0_00040102',2,ns_s,ns_sz))
    while len(dir_entries)%4: dir_entries.append(empty_de())

    dir_s,_=alloc(b''.join(dir_entries))

    # FAT
    fat_idx=len(sectors)
    n_fat=max(1,(len(fat)+1+127)//128)
    fat_full=fat+[FREESECT]*(n_fat*128-len(fat))
    for i in range(n_fat):
        idx=fat_idx+i
        if idx<len(fat_full): fat_full[idx]=FATSECT
        else: fat_full.append(FATSECT)
    for i in range(n_fat):
        sectors.append(struct.pack('<128I',*fat_full[i*128:(i+1)*128]))

    hdr=bytearray(512)
    hdr[0:8]=b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
    struct.pack_into('<H',hdr,24,0x3E); struct.pack_into('<H',hdr,26,3)
    struct.pack_into('<H',hdr,28,0xFFFE)
    struct.pack_into('<H',hdr,30,9); struct.pack_into('<H',hdr,32,6)
    struct.pack_into('<I',hdr,40,0); struct.pack_into('<I',hdr,44,n_fat)
    struct.pack_into('<I',hdr,48,dir_s)
    struct.pack_into('<I',hdr,52,0); struct.pack_into('<I',hdr,56,0x1000)
    struct.pack_into('<I',hdr,60,ENDOFCHAIN); struct.pack_into('<I',hdr,64,0)
    struct.pack_into('<I',hdr,68,ENDOFCHAIN); struct.pack_into('<I',hdr,72,0)
    for i in range(min(n_fat,109)):
        struct.pack_into('<I',hdr,76+i*4,fat_idx+i)
    for i in range(n_fat,109):
        struct.pack_into('<I',hdr,76+i*4,FREESECT)

    return bytes(hdr)+b''.join(sectors)


def main():
    base_dir  = get_exe_dir()
    zip_files = glob.glob(os.path.join(base_dir,'*.zip'))

    print('='*52)
    print('  ZIP → MSG Converter')
    print('='*52)

    if not zip_files:
        print(f'\n[오류] 같은 폴더에 .zip 파일이 없습니다.')
        print(f'  폴더: {base_dir}')
        input('\nEnter 키를 눌러 종료...')
        sys.exit(1)

    print(f'\n  {len(zip_files)}개 파일 발견\n')

    ok=fail=0
    for zip_path in zip_files:
        base=os.path.splitext(zip_path)[0]
        msg_path=base+'.msg'
        fname=os.path.basename(zip_path)
        try:
            out=build_zip_msg(zip_path)
            with open(msg_path,'wb') as f: f.write(out)
            print(f'  ✓  {fname}  →  {os.path.basename(msg_path)}')
            ok+=1
        except Exception as e:
            print(f'  ✗  {fname}  ({e})')
            fail+=1

    print(f'\n  완료: 성공 {ok}개 / 실패 {fail}개')
    input('\nEnter 키를 눌러 종료...')


if __name__=='__main__':
    main()
