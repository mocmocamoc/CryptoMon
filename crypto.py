import xlwings as xw
import requests
import json
import time
try:
    from bottle import get, run
except:
    pass
import threading
import os
import hmac
import hashlib
import csv
from crypto_priv import HOLDINGS, CHANGELLY_KEY, CHANGELLY_SECRET

LIMIT=40

class CMCClient:
    data = None
    fx = None
    dataTime = 0
    fxTime = 0
    
    def getData(this):
        if time.time() - this.dataTime >= 6:
            this.dataTime = time.time()
            try:
                resp = requests.get("https://api.coinmarketcap.com/v1/ticker/?limit=%s" % LIMIT)
                #resp = requests.get("https://api.coinmarketcap.com/v1/ticker/")
                if resp.ok:
                    this.data = resp.json()
            except:
                pass
        return this.data
        
    def getFX(this, base="USD", target="GBP"):
        if time.time() - this.fxTime >= (60 * 60):
            this.fxTime = time.time()
            try:
                resp = requests.get("http://api.fixer.io/latest?base=%s" % base)
                if resp.ok:
                    this.fx = resp.json()["rates"][target]
            except:
                pass
        return this.fx

client=CMCClient()

@xw.func
def getFX(base="USD", target="GBP"):
    return client.getFX(base, target)

@xw.func
def getCaps():
    l = client.getData()
    return [(x["symbol"], float(x["market_cap_usd"])) for x in l if x["market_cap_usd"]]

@xw.func
def getPrice(id="BTC"):
    l = client.getData()
    return float([i for i in l if i["symbol"] == id][0]["price_usd"])
    
@xw.func
def getBitty():
    return requests.get("https://bittylicious.com/api/v1/quote/BTC/GB/GBP/BANK/1").json()["totalPrice"]

def toNum(s):
    if not s:
        return 0
    else:
        try:
            return int(s)
        except ValueError:
            return float(s)    
    
class CMCReport:
    fileHist = os.path.join(os.path.dirname(os.path.realpath(__file__)), "hist.csv")
    
    def loadHist(self):
        if os.path.exists(self.fileHist):
            with open(self.fileHist, "r") as f:
                r = csv.DictReader(f)
                tmpHist = [i for i in r]
            self.hist = dict([(int(t["Time"]),dict([(i, toNum(j)) for i, j in list(t.items()) if i != "Time"])) for t in tmpHist])
        else:
            self.hist = {}
            
    def saveHist(self):
        keys = [k for k,v in sorted(list(self.hist[max(self.hist)].items()), reverse=True, key=lambda x_y: x_y[1])]
        #self.hist = dict([(k,v) for k, v in self.hist.items() if k > time.time() - (24 * 60 * 60)])
        saveHist = sorted([dict([(i,j) for i,j in list(v.items()) if i in keys] + [("Time", k)]) for k, v in list(self.hist.items())], reverse=True, key=lambda x: x["Time"])
        with open(self.fileHist, "w", newline='') as f:
            w = csv.DictWriter(f, ["Time"] + keys)
            w.writeheader()
            w.writerows(saveHist)
    
    def __init__(self):
        self.loadHist()
        self.useBS = False
        self.useCF = False
        try:
            resp = requests.get("https://webapi.coinfloor.co.uk:8090/bist/XBT/GBP/order_book/")
            if resp.ok:
                self.useCF = True
        except:
            pass
    
    def getSummary(self):
        sleepTime = 6
        while True:
            fx = getFX("USD", "GBP")
            data = client.getData()
            prices = dict([(i["symbol"], float(i["price_usd"])) for i in data if i["price_usd"] is not None])
            det = {"Total": 0}
            for id, qty in list(HOLDINGS.items()):
                val = prices[id] * qty * fx
                det[id] = int(round(val))
                det["Total"] += val
            det["Total"] = int(round(det["Total"]))
            if not self.hist or det["Total"] != self.hist[max(self.hist)]["Total"]:
                #sleepTime = 300
                equiv = self.getEquivalent("BTC", ignoreLimit=True)
                for id, qty in list(HOLDINGS.items()):
                    if id not in equiv:
                        equiv[id] = ("NA", (prices[id] * qty) / prices["BTC"])
                equivSum = sum([v[1] for k,v in list(equiv.items())])
                det["Equiv"] = round(equivSum, 2)
                if self.useBS:
                    equivSumNoXRP = sum([v[1] for k,v in list(equiv.items()) if k != "XRP"])
                    real = self.getSalePrice(equivSumNoXRP, "BTC", "USD") + self.getSalePrice(HOLDINGS.get("XRP", 0), "XRP", "USD") 
                    det["Real"] = int(round(real * fx))
                elif self.useCF:
                    real = self.getSalePriceCF(equivSum)
                    det["Real"] = int(round(real))
                else:
                    det["Real"] = int(round(equivSum * prices["BTC"] * fx))
                self.hist[int(time.time())] = det
                self.hist = dict([(k,v) for k, v in list(self.hist.items()) if k > time.time() - (24 * 60 * 60)])
                self.saveHist()
                print("{:s}: Updated: {:,d}".format(time.strftime("%H:%M:%S"), det["Total"]))
            time.sleep(sleepTime)
           
    def getSalePrice(self, coins, frm="BTC", to="USD"):
        price = 0
        resp = requests.get("https://www.bitstamp.net/api/v2/order_book/%s%s" % (frm.lower(), to.lower()))
        if resp.ok:
            bids = resp.json()["bids"]
            bids = [(float(b), float(a)) for b, a in bids]
            for bid, amount in sorted(bids, reverse=True):
                if amount >= coins:
                    price += coins * bid
                    break
                else:
                    price += amount * bid
                    coins -= amount
        else:
            print(resp.url, resp.text)
        return (price - 10) * 0.9962
    
    def getSalePriceCF(self, coins):
        price = 0
        try:
            resp=requests.get("https://webapi.coinfloor.co.uk:8090/bist/XBT/GBP/order_book/")
        except:
            return 0
        if resp.ok:
            bids = resp.json()["bids"]
            bids = [(float(b), float(a)) for b, a in bids]
            for bid, amount in sorted(bids, reverse=True):
                if amount >= coins:
                    price += coins * bid
                    break
                else:
                    price += amount * bid
                    coins -= amount
        else:
            print(resp.url, resp.text)
        return price * 0.9975

    def getEquivalent(self, target="BTC", ignoreLimit=False):
        equivSS = {}
        equivCY = {}
        equiv = {}
        for k in HOLDINGS:
            if k.lower() == target.lower():
                equiv[k] = ("NA", HOLDINGS[k])
                continue
            pair = "%s_%s" % (k.lower(), target.lower())
            resp = requests.get("https://shapeshift.io/marketinfo/" + pair)
            if resp.ok and "rate" in resp.json() and (ignoreLimit or ("limit" in resp.json() and float(resp.json()["limit"]) >= HOLDINGS[k])):
                equivSS[k] = float(resp.json()["rate"]) * HOLDINGS[k]
            h = hmac.new(CHANGELLY_SECRET, digestmod=hashlib.sha512)
            x = {'id': 1, 'jsonrpc': '2.0', 'method': 'getExchangeAmount', 'params': {'amount': HOLDINGS[k], 'from': k.lower() if k.lower() != "bch" else "bcc", 'to': target.lower()}}
            h.update(json.dumps(x))
            resp = requests.post("https://api.changelly.com", data=json.dumps(x), headers={"api-key": CHANGELLY_KEY, "sign": h.hexdigest(), "content-type": "application/json"})
            if resp.ok and "result" in resp.json():
                equivCY[k] = float(resp.json()["result"])
            if k in equivCY and equivCY[k] > equivSS.get(k, 0):
                equiv[k] = ("CY", equivCY[k])
            elif k in equivSS and equivSS[k] > equivCY.get(k, 0):
                equiv[k] = ("SS", equivSS[k])
            elif k in equivSS and k in equivCY and equivSS[k] == equivCY[k]:
                equiv[k] = ("SS", equivSS[k])
            else:
                print("No exchange found for %s!" % k)
        return equiv
    
    def getExitPlan(self, target="BTC"):
        for ignoreLimit in (True, False):
            yield "ignoreLimit: <b>%s</b><br><br>" % ignoreLimit
            equiv = self.getEquivalent(target, ignoreLimit)
            fx = getFX("USD", "GBP")
            for keepXRP in [False]:
                yield "keepXRP: <b>%s</b><br><br>" % keepXRP
                for k,v in sorted(equiv.items()):
                    if v[0] == "NA" or (k == "XRP" and self.useBS and keepXRP):
                        yield "{:s}: Keep {:.8f} {:s}".format(k, HOLDINGS[k], k)
                    elif v[0] == "CY":
                        yield "{:s}: Changelly {:.8f} {:s} to {:,.8f} {:s}".format(k, HOLDINGS[k], k, v[1], target)
                    elif v[0] == "SS":
                        yield "{:s}: Shapeshift {:.8f} {:s} to {:,.8f} {:s}".format(k, HOLDINGS[k], k, v[1], target)
                    yield "<br>"
                yield "<br>"
                equivSum = sum([v[1] for k, v in list(equiv.items()) if not (k == "XRP" and self.useBS and keepXRP)])
                yield "Total {:s}: {:.8f}<br>".format(target, equivSum)
                if self.useBS:
                    if keepXRP:
                        yield "Total XRP: {:.8f}<br>".format(HOLDINGS["XRP"])
                    yield "<br>"
                    sellTarg = self.getSalePrice(equivSum, target, "USD")
                    sellXRP = self.getSalePrice(HOLDINGS["XRP"], "XRP", "USD") if keepXRP else 0
                    yield "Sell {:s} for {:,.2f} USD<br>".format(target, sellTarg)
                    if keepXRP:
                        yield "Sell XRP for {:,.2f} USD<br>".format(sellXRP)
                    yield "<br>"
                    yield "Total USD {:,.2f}<br>".format(sellTarg + sellXRP)
                    yield "Total GBP {:,.2f}".format((sellTarg + sellXRP) * fx)
                yield "<hr>"

    def getHistory(self):
        if not self.hist:
            yield("No data")
            return
        totals = [d["Total"] for k, d in sorted(self.hist.items())]
        symbol = ""
        if len(totals) > 1:
            if totals[-1] > totals[-2]:
                symbol = chr(0x25b2)
            elif totals[-1] < totals[-2]:
                symbol = chr(0x25bc)
            else:
                symbol = chr(0x25cf)
        if totals[-1] == max(totals):
            symbol = chr(0x2605)
        yield("<html>\n")
        yield("<head><title>{:,.0f} {:s}</title><meta http-equiv='Content-Type' content='text/html; charset=UTF-8'><meta http-equiv='refresh' content='60'></head>\n".format(totals[-1], symbol))
        yield("<body><style>table, th, td { border: 1px solid black; border-collapse: collapse; padding: 3px;} td { text-align: right;}</style>")
        yield("<table>")
        keys = [k for k,v in sorted(list(self.hist[max(self.hist)].items()), reverse=True, key=lambda x_y1: x_y1[1])]
        yield("<tr>")
        yield("<th>Time</th>")
        yield("".join(["<th>%s</th>" % k for k in keys]))
        yield("</tr>")
        rows = []
        dprev = None
        #toShow = sorted([(k,v) for k, v in self.hist.items() if k > time.time() - (24 * 60 * 60)])
        for t, d in sorted(self.hist.items()):
            row = "<tr><td>%s</td>" % time.strftime("%H:%M:%S", time.localtime(t))
            for k in keys:
                if k not in d:
                    d[k] = 0
                col = "black"
                if dprev is not None:
                    if d[k] > dprev[k]:
                        col = "green"
                    elif d[k] < dprev[k]:
                        col = "red"
                strTD = "<td style='color:{:s}'>{:,.%if}</td>" % (2 if d[k] > 0 and d[k] < 100 else 0)
                row += strTD.format("blue" if d[k] == max([v.get(k, 0) for v in list(self.hist.values())]) else col, d[k])
            dprev = d
            row += "</tr>"
            rows = [row] + rows
        yield("\n".join(rows))
        yield("</table>")
        yield("\n</body></html>")
        
    def run(self):
        get("/history")(self.getHistory)
        get("/exit")(self.getExitPlan)
        run(host='0.0.0.0', port=8080, server='paste', debug=True)
            
def main():
    r = CMCReport()
    thread = threading.Thread(name="gs", target=r.getSummary)
    thread.daemon = True
    thread.start()
    r.run()
    print("Exiting")

if __name__ == "__main__":
    main()