import { useEffect, useRef, useState } from 'react';

/** The editor's notice: an error stays until replaced, a warning clears itself after three seconds. */
export function useNotice() {
  const [notice, setNoticeState] = useState<string | null>(null);
  const errorNotice = useRef<string | null>(null);
  const [noticeTone, setNoticeTone] = useState<'danger' | 'warning'>('danger');
  const [noticeRevision, setNoticeRevision] = useState(0);
  function setNotice(message: string | null) {
    errorNotice.current = message;
    setNoticeTone('danger');
    setNoticeState(message);
  }
  function warn(message: string) {
    errorNotice.current = null;
    setNoticeTone('warning');
    setNoticeState(message);
    setNoticeRevision((value) => value + 1);
  }
  useEffect(() => {
    if (notice && noticeTone === 'warning') {
      const timer = setTimeout(() => setNoticeState(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [notice, noticeTone, noticeRevision]);
  return { notice, noticeTone, noticeRevision, errorNotice, setNotice, warn };
}
