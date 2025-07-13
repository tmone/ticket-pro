"use client";

import { useEffect } from "react";
import { useSearchParams } from "next/navigation";
import { getRedirectUrl } from "@/config/checkin-redirects";

export function RedirectHandler() {
  const searchParams = useSearchParams();
  
  useEffect(() => {
    const code = searchParams.get('code');
    if (code) {
      const redirectUrl = getRedirectUrl(code);
      if (redirectUrl) {
        window.location.href = redirectUrl;
      }
    }
  }, [searchParams]);
  
  return null;
}