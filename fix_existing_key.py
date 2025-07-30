#!/usr/bin/env python3
"""
기존 서비스 키에서 제어 문자를 제거하여 올바른 base64 값을 생성하는 스크립트
"""

import base64
import json
import re


def fix_service_key_base64(corrupted_base64):
    """
    손상된 base64에서 디코딩하여 제어 문자를 제거하고 다시 인코딩
    """
    try:
        print("🔍 원본 base64 길이:", len(corrupted_base64))

        # 1. base64 디코딩
        cleaned_b64 = "".join(corrupted_base64.split())
        decoded_bytes = base64.b64decode(cleaned_b64)
        decoded_text = decoded_bytes.decode("utf-8", errors="replace")

        print("🔍 디코딩된 텍스트 길이:", len(decoded_text))

        # 2. 제어 문자 제거 및 수정
        fixed_text = decoded_text

        # 특정 문제 수정: PRIVATE^FKEY → PRIVATE KEY
        fixed_text = fixed_text.replace("PRIVATE\x06KEY", "PRIVATE KEY")
        fixed_text = fixed_text.replace("PRIVATE^FKEY", "PRIVATE KEY")

        # 모든 제어 문자 제거 (0x00-0x1F, 0x7F 제외하고 유지할 것들: \t(0x09), \n(0x0A), \r(0x0D))
        fixed_text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", fixed_text)

        print("🔧 수정된 텍스트 길이:", len(fixed_text))

        # 3. JSON 유효성 검사
        try:
            json_data = json.loads(fixed_text)
            print("✅ JSON 유효성 검사 통과")

            # 4. 다시 base64 인코딩
            fixed_json = json.dumps(json_data, separators=(",", ":"))  # 압축된 JSON
            new_base64 = base64.b64encode(fixed_json.encode("utf-8")).decode("ascii")

            print("🎉 새로운 base64 길이:", len(new_base64))
            return new_base64, json_data

        except json.JSONDecodeError as e:
            print(f"❌ JSON 파싱 실패: {e}")
            print("🔍 문제가 있는 부분:")
            error_pos = getattr(e, "pos", 0)
            start = max(0, error_pos - 50)
            end = min(len(fixed_text), error_pos + 50)
            print(repr(fixed_text[start:end]))
            return None, None

    except Exception as e:
        print(f"❌ 처리 중 오류: {e}")
        return None, None


def main():
    print("🔧 서비스 키 제어 문자 수정 도구")
    print("=" * 50)

    # GitHub Secrets에서 복사한 base64 값을 여기에 붙여넣으세요
    print("📋 GitHub Secrets의 SHEETS_SERVICE_ACCOUNT_KEY 값을 붙여넣으세요:")
    print("(Ctrl+C로 종료)")

    try:
        corrupted_base64 = input().strip()

        if not corrupted_base64:
            print("❌ 입력이 비어있습니다.")
            return

        print("\n🔄 처리 중...")
        fixed_base64, json_data = fix_service_key_base64(corrupted_base64)

        if fixed_base64:
            print("\n" + "=" * 50)
            print("🎉 수정 완료!")
            print("=" * 50)
            print("\n📝 GitHub Secrets에 업데이트할 새로운 값:")
            print("-" * 50)
            print(fixed_base64)
            print("-" * 50)

            print(f"\n📊 서비스 계정 정보:")
            print(f"   프로젝트 ID: {json_data.get('project_id', 'N/A')}")
            print(f"   클라이언트 이메일: {json_data.get('client_email', 'N/A')}")
            print(f"   키 ID: {json_data.get('private_key_id', 'N/A')[:16]}...")

            # 파일로도 저장
            with open("fixed_service_key.json", "w") as f:
                json.dump(json_data, f, indent=2)
            print(f"\n💾 수정된 JSON이 'fixed_service_key.json'에 저장되었습니다.")

            print(f"\n🔄 다음 단계:")
            print(f"1. 위의 base64 값을 복사")
            print(f"2. GitHub → Settings → Secrets → SHEETS_SERVICE_ACCOUNT_KEY 업데이트")
            print(f"3. DRIVE_SERVICE_ACCOUNT_KEY도 동일한 방식으로 처리")

        else:
            print("❌ 키 수정에 실패했습니다.")

    except KeyboardInterrupt:
        print("\n\n👋 프로그램을 종료합니다.")
    except Exception as e:
        print(f"\n❌ 예상치 못한 오류: {e}")


if __name__ == "__main__":
    main()
