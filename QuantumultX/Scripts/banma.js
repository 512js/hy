/*
微信公众号：ios黑科技
官方网站：s7aa.cn

[rewrite_local]

#斑马海报解锁终身会员
^^https:\/\/zebra\.maka\.im\/api\/user\/info url script-response-body banma.js


[mitm]
hostname = zebra.maka.im
*/

let obj = JSON.parse($response.body);
obj = {
  "data": {
    "business": {
      "wool": 5201314,
      "is_lifelong_vip": true
    },
    "profile": {
      "uid": 700000000,
      "nickname": "Hy",
      "avatar": "https://i.loli.net/2020/07/04/NBrHSoDL1TEfdjX.png"
    }
  },
  "error": "",
  "resultCode": 0
}

$done({body: JSON.stringify(obj)});
