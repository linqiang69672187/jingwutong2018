/**
 *  卷轴控件
 *  @author:dorseyCh
 *  @date:2018-09-13
***/
(function(win){

    var Scroll = function(option){
        this.option = option;
    };

    Scroll.prototype = {
        //初始化
        init : function(){
            this.oldData = 900;
            this.newData = 800;
            this.arrOld = [0,0,0,0,0];
            this.arrNew = [0,0,0,0,0];
            this.direction = [];
            this.styleReturn();
        },
        //数据生成器 ———— 配合后台则需要异步接收到数据
        dataGet : function () {
            return Math.floor(Math.random() * 99999);
        },
        //数据处理
        dataSolve : function () {

            var num = this.option.num;
            var newData = this.dataGet(),
                oldData = this.newData;

            this.newData = newData;
            var arrOld = [],arrNew = [];
            for(var i = 0 ; i < num ; i ++) {
                arrOld[num - 1 - i] = (parseInt(oldData) % 10);
                // if(parseInt(newData) % 10 !== parseInt(oldData) % 10){
                arrNew[num - 1 - i] = (parseInt(newData) % 10);
                // }
                newData = Math.floor(newData / 10);
                oldData = Math.floor(oldData / 10);
            }
            this.arrOld = arrOld;
            this.arrNew = arrNew;
            this.oldData = this.newData;
            
        },
        //数据展示
        dataDisplay : function (){
            var _self = this;
            var $li = _self.option.dom.find("li");
            _self.direction = [];

            for(var i = 0 ; i < _self.option.num ; i ++){

                //新数据与旧数据对比，较大时，div-hide瞬移到div-on的下方，反之则上方，同时返回一个数组表示接下来li要移动的方向
                // if(_self.arrNew[i] !== '') {
                    _self.arrOld[i] > _self.arrNew[i] ? 
                    ($li.eq(i).find('.s-hide').attr("class","s-hide s-top"), _self.direction.push('bottom')) : 
                    ($li.eq(i).find('.s-hide').attr("class","s-hide s-bottom"), _self.direction.push('top'));
                // }
                // else{
                //     _self.direction.push('');
                // }
                $li.eq(i).find('.s-on').html(_self.arrOld[i]);
                $li.eq(i).find('.s-hide').html(_self.arrNew[i]);
            }

        },
        //添加动画
        animate : function () {
            var _self = this;           
            var $li = this.option.dom.find('li');
            
            $li.each(function(i){
                $li.eq(i).css("transition", 0.5 * ($li.length - i) + "s").attr("class","s-" + _self.direction[i]);
            });
        },
        //样式归原
        styleReturn : function () {
            var _self = this;
            var $li = _self.option.dom.find("li");
            
            for(var i = 0 ; i < _self.option.num ; i ++){
                $li.eq(i).find('.s-on').html(_self.arrNew[i]);
                $li.eq(i).find('.s-hide').html(_self.arrOld[i]);
            }

            $li.css("transition","").attr('class','');
            $li.find(".s-on").attr("class","s-on"),
            $li.find(".s-hide").attr("class","s-hide s-bottom");
            
        },
        run : function () {
            var _self = this;
            _self.init();
            
            setInterval(function(){
                //这里暂时想不到更好的办法，先勉强这样
                _self.styleReturn();    //数据还原
                setTimeout(function(){
                    _self.dataSolve();      //数据处理
                    _self.dataDisplay();    //数据展示
                    _self.animate();
                },100);
            },_self.option.delay);
        }
    }

    win.Scroll = Scroll;

})(window);