using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WeighingSystem
{
    class WebFax
    {
        /*
         * 팝빌 팩스 API DotNet SDK Example
         * 
         * - DotNet C# SDK 연동환경 설정방법 안내 : [개발가이드] - https://docs.popbill.com/fax/tutorial/dotnet#csharp
         * - 업데이트 일자 :  2020-01-21
         * - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991~2
         * - 연동 기술지원 이메일 : code@linkhub.co.kr
         * 
         * <테스트 연동개발 준비사항>
         * 1) 29, 32 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
         *    링크허브 가입시 메일로 발급받은 인증정보로 변경합니다.
         * 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
         */

        #region Properties
        
        /// <summary>
        /// 연동신청시 발급받은 링크아이디
        /// </summary>
        private string LinkID = "LST";

        /// <summary>
        /// 연동신청시 발급받은 비밀키
        /// </summary>
        //private string SecretKey = "Ljlyy4McNWmMIO+p77pE8QgDfe/AM4lQAVcGLvG30Jc=";
        private string SecretKey = "M0XgPfNhJ0+c0KKTgjtgme9oT445+A0BYQiKZqzYAq0=";
        /// <summary>
        /// 팝빌회원 사업자번호
        /// </summary>
        //private string corpNum = "6062884137";
        private string corpNum = "6068188502";

        /// <summary>
        /// 팝빌회원 아이디
        /// </summary>
        //private string userID = "steelnet";
        private string userID = "daejist";

        /// <summary>
        /// 팩스 서비스 객체
        /// </summary>
        private Popbill.Fax.FaxService _faxService;

        /// <summary>
        /// 연동환경 설정값, true - 개발용(테스트베드), false - 상업용(실서비스)
        /// </summary>
        public bool IsTest
        {
            get { return _isTest; }
            set { _isTest = value; }
        }
        private bool _isTest = true;

        #endregion

        #region Constructor

        /// <summary>
        /// WebFax 생성자
        /// </summary>
        /// <param name="isTest"></param>
        public WebFax(bool isTest)
        {
            // 팩스 서비스 객체 초기화
            _faxService = new Popbill.Fax.FaxService(LinkID, SecretKey);

            // 연동환경 설정값, true - 개발용(테스트베드), false - 상업용(실서비스)
            _faxService.IsTest = isTest;

            // 발급된 토큰에 대한 IP 제한기능 사용여부, 권장(True)
            _faxService.IPRestrictOnOff = true;
        }

        #endregion

        #region Methods

        /// <summary>
        /// WebFax 를 전송함
        /// </summary>
        /// <param name="senderNum">발신번호</param>
        /// <param name="receiverNum">수신번호</param>
        /// <param name="receiverName">수신자명</param>
        /// <param name="fileName">파일명</param>
        /// <param name="title">팩스제목</param>
        /// <param name="result">out result</param>
        /// <returns></returns>
        public bool SendFax(string senderNum, string receiverNum, string receiverName, string fileName, string title, out string result)
        {
            // 광고팩스 전송여부
            bool adsYN = false;

            // 전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
            // 최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
            String requestNum = "";

            DateTime? reserveDT = null;

            try
            {
                string receiptNum = _faxService.SendFAX
                (
                    corpNum, senderNum, receiverNum, receiverName,
                    fileName, reserveDT, userID, adsYN, title, requestNum
                );

                // 접수번호
                result = receiptNum;

                return true;
            }
            catch (Popbill.PopbillException ex)
            {
                // 오류내용
                result = string.Concat
                (
                    "응답코드(code) : ", ex.code.ToString(), "\r\n",
                    "응답메시지(message) : ", ex.Message
                );

                return false;
            }
        }

        /// <summary>
        /// 팩스전송요청시 발급받은 접수번호(receiptNum)로 전송결과를 확인
        /// </summary>
        /// <param name="receiptNum">접수번호</param>
        /// <returns></returns>
        public string GetFaxResult(string receiptNum)
        {
            try
            {
                List<Popbill.Fax.FaxResult> ResultList = _faxService.GetFaxResult(corpNum, receiptNum);

                StringBuilder sb = new StringBuilder();
                
                for (int i = 0; i < ResultList.Count; i++)
                {
                    string fileList = string.Empty;
                    for (int j = 0; j < ResultList[i].fileNames.Count(); j++)
                    {
                        if (j == ResultList[i].fileNames.Count() - 1)
                            fileList += ResultList[i].fileNames[j].ToString();
                        else
                            fileList += ResultList[i].fileNames[j].ToString() + ",";
                    }

                    sb.AppendFormat("state(전송상태 코드) : {0} \n\n", ResultList[i].state);
                    sb.AppendFormat("result(전송결과 코드) : {0} \n\n", ResultList[i].result);
                    sb.AppendFormat("sendNum(발신번호) : {0} \n\n", ResultList[i].sendNum);
                    sb.AppendFormat("senderName(발신자명) : {0} \n\n", ResultList[i].senderName);
                    sb.AppendFormat("receiveNum(수신번호) : {0} \n\n", ResultList[i].receiveNum);
                    sb.AppendFormat("receiveName(수신자명) : {0} \n\n", ResultList[i].receiveName);
                    sb.AppendFormat("title(팩스제목) : {0} \n\n", ResultList[i].title);
                    sb.AppendFormat("sendPageCnt(전체 페이지수) : {0} \n\n", ResultList[i].sendPageCnt);
                    sb.AppendFormat("successPageCnt(성공 페이지수) : {0} \n\n", ResultList[i].successPageCnt);
                    sb.AppendFormat("failPageCnt(실패 페이지수) : {0} \n\n", ResultList[i].failPageCnt);
                    sb.AppendFormat("refundPageCnt(환불 페이지수) : {0} \n\n", ResultList[i].refundPageCnt);
                    sb.AppendFormat("cancelPageCnt(취소 페이지수) : {0} \n\n", ResultList[i].cancelPageCnt);
                    sb.AppendFormat("reserveDT(예약시간) : {0} \n\n", ResultList[i].reserveDT);
                    sb.AppendFormat("receiptDT(접수시간) : {0} \n\n", ResultList[i].receiptDT);
                    sb.AppendFormat("sendDT(발송시간) : {0} \n\n", ResultList[i].sendDT);
                    sb.AppendFormat("resultDT(전송결과 수신시간) : {0} \n\n", ResultList[i].resultDT);
                    sb.AppendFormat("fileNames(전송 파일명 리스트) : {0} \n\n", fileList);
                    sb.AppendFormat("receiptNum(접수번호) : {0} \n\n", ResultList[i].receiptNum);
                    sb.AppendFormat("requestNum(요청번호) : {0} \n\n", ResultList[i].requestNum);
                    sb.AppendFormat("chargePageCnt(과금 페이지수) : {0} \n\n", ResultList[i].chargePageCnt);
                    sb.AppendFormat("tiffFileSize(변환파일용량(단위:byte)) : {0} \n\n", ResultList[i].tiffFileSize);
                }

                return sb.ToString();
            }
            catch (Popbill.PopbillException ex)
            {
                // 오류내용
                string error = string.Concat
                (
                    "응답코드(code) : ", ex.code.ToString(), "\r\n",
                    "응답메시지(message) : ", ex.Message
                );

                return error;
            }
        }

        #endregion

        #region Implementations

        #endregion
    }
}
