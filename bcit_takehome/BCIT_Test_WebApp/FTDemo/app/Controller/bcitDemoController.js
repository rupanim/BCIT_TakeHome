

var controllerId = 'bcitDemoController';

    angular.module('app').controller(controllerId, Demo);

    Demo.$inject = ['$scope', '$location', '$anchorScroll',  '$sce', '$http'];

function Demo($scope, $location, $anchorScroll, $sce, $http) {     

    $scope.fileLoc = "";

    function reset() {
        $scope.validationErrMsg = "";
        $scope.isValidationError = false;
        $scope.isParsed = false;
        $scope.isParsing = false;
        
        $scope.outputHtml = "";
        $scope.parsingErrMsg = "";
        $scope.isParsingError = false;
        $scope.parsedHtmlLoc = "";
    }    

    function validate(filePath) {
        filePath = filePath.trim();

        if (filePath == "") {
            $scope.validationErrMsg = "File Path cannot be empty.";
            $scope.isValidationError = true;
            return false;
        }      

        if (filePath.length == 0 || filePath.substr(filePath.lastIndexOf('.') + 1) != "docx") {
            $scope.validationErrMsg = "Invalid file extension/path.";
            $scope.isValidationError = true;
            return false;
        }

        return true;
    }    

    $scope.parseNow = function () {

        reset();

        if (!validate($scope.fileLoc)) {
            return;
        }

        var url = "/api/parserapi/parser/get?file=" + $scope.fileLoc, data = "";

        $scope.isParsing = true;

        $http.get(url, $scope.fileLoc).then(function (response) {
            $scope.isParsing = false;
            var t = response.data;
            if (t == "-1") {
                $scope.validationErrMsg = "File does not exists.";
                $scope.isValidationError = true;
            } else {
                $scope.outputHtml = $sce.trustAsHtml(t);                
            }
            $scope.parsedHtmlLoc = $scope.fileLoc.substr(0, $scope.fileLoc.lastIndexOf('.')) + ".htm";
            $scope.isParsed = true;
        }, function (response) {
                $scope.isParsing = false;
                $scope.parsingErrMsg = response.data;
                $scope.isParsingError = true;
                $scope.isParsed = true;
        });
            
    };

    $scope.gotoAnchor = function (x) {
        var newHash = 'anchor' + x;
        if ($location.hash() !== newHash) {              
            $location.hash('anchor' + x);
        } else {               
            $anchorScroll();
        }
    };

    reset();

    }


  